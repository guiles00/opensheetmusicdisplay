import { expect } from "chai";
import {
    IVirtualSystemDescriptor,
    SystemVirtualizationController
} from "../../src/OpenSheetMusicDisplay/SystemVirtualizationController";

describe("SystemVirtualizationController offscreen drawing pace", () => {
    let container: HTMLDivElement;
    let scrollElement: HTMLDivElement;
    let svg: SVGSVGElement;
    let frames: FrameRequestCallback[];
    let realRequestAnimationFrame: typeof window.requestAnimationFrame;
    let realCancelAnimationFrame: typeof window.cancelAnimationFrame;
    let realPerformanceNow: () => number;
    let clockMs: number;

    const runFrame: () => void = (): void => {
        const pending: FrameRequestCallback[] = frames;
        frames = [];
        for (const frame of pending) {
            frame(clockMs);
        }
    };

    beforeEach(() => {
        container = document.createElement("div");
        scrollElement = document.createElement("div");
        svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
        document.body.appendChild(scrollElement);
        scrollElement.appendChild(container);
        container.appendChild(svg);
        scrollElement.getBoundingClientRect = (): DOMRect => ({
            x: 0, y: 0, top: 0, right: 100, bottom: 100,
            left: 0, width: 100, height: 100, toJSON: (): object => ({})
        });
        svg.getScreenCTM = (): DOMMatrix => new DOMMatrix();

        frames = [];
        clockMs = 1000;
        realRequestAnimationFrame = window.requestAnimationFrame;
        realCancelAnimationFrame = window.cancelAnimationFrame;
        realPerformanceNow = performance.now.bind(performance);
        window.requestAnimationFrame = ((callback: FrameRequestCallback): number => {
            frames.push(callback);
            return frames.length;
        }) as typeof window.requestAnimationFrame;
        window.cancelAnimationFrame = ((): void => undefined) as typeof window.cancelAnimationFrame;
        performance.now = (): number => clockMs;
    });

    afterEach(() => {
        window.requestAnimationFrame = realRequestAnimationFrame;
        window.cancelAnimationFrame = realCancelAnimationFrame;
        performance.now = realPerformanceNow;
        scrollElement.remove();
    });

    /** One offscreen system per key, each taking `drawCostMs` of the fake clock to draw. */
    const setup: (drawCostMs: number) => { controller: SystemVirtualizationController, drawnKeys: string[] } =
        (drawCostMs: number): { controller: SystemVirtualizationController, drawnKeys: string[] } => {
            const descriptors: IVirtualSystemDescriptor[] = [];
            for (let index: number = 0; index < 6; index++) {
                descriptors.push({ key: `1:${index}`, svg, top: 200 + index * 50, bottom: 240 + index * 50 });
            }
            const drawnKeys: string[] = [];
            const controller: SystemVirtualizationController = new SystemVirtualizationController(container);
            controller.enable({ scrollElement, overscanViewports: 4, idleMaterializationBudgetMs: 10 });
            controller.configureExpectedSystems(descriptors, (keys: string[]): SVGGElement[] => {
                clockMs += drawCostMs * keys.length;
                const groups: SVGGElement[] = [];
                for (const key of keys) {
                    drawnKeys.push(key);
                    const group: SVGGElement = document.createElementNS("http://www.w3.org/2000/svg", "g");
                    group.classList.add("osmd-system");
                    group.dataset.osmdSystemKey = key;
                    const descriptor: IVirtualSystemDescriptor = descriptors.find(candidate => candidate.key === key);
                    group.dataset.osmdSystemTop = descriptor.top.toString();
                    group.dataset.osmdSystemBottom = descriptor.bottom.toString();
                    svg.appendChild(group);
                    groups.push(group);
                }
                return groups;
            });
            return { controller, drawnKeys };
        };

    it("keeps drawing every frame while systems are cheap", () => {
        const { drawnKeys } = setup(1);
        const drawnAfterConfigure: number = drawnKeys.length;
        runFrame();
        runFrame();
        expect(drawnKeys.length).to.be.greaterThan(drawnAfterConfigure);
    });

    it("skips the frames a slow system overran instead of drawing in every one", () => {
        const { drawnKeys } = setup(40);
        const drawnAfterConfigure: number = drawnKeys.length;
        runFrame();
        const afterFirstFrame: number = drawnKeys.length;
        expect(afterFirstFrame).to.equal(drawnAfterConfigure + 1);
        runFrame();
        runFrame();
        runFrame();
        expect(drawnKeys.length).to.equal(afterFirstFrame);
        runFrame();
        expect(drawnKeys.length).to.equal(afterFirstFrame + 1);
    });
});
