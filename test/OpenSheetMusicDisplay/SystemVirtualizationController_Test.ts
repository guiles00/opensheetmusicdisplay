import { expect } from "chai";
import { SystemVirtualizationController } from "../../src/OpenSheetMusicDisplay/SystemVirtualizationController";

describe("SystemVirtualizationController", () => {
    it("detaches distant systems and restores the same stateful SVG nodes", () => {
        const container: HTMLDivElement = document.createElement("div");
        const scrollElement: HTMLDivElement = document.createElement("div");
        const svg: SVGSVGElement = document.createElementNS("http://www.w3.org/2000/svg", "svg");
        document.body.appendChild(scrollElement);
        scrollElement.appendChild(container);
        container.appendChild(svg);

        let scrollTop: number = 0;
        scrollElement.getBoundingClientRect = (): DOMRect => ({
            x: 0, y: 0, top: 0, right: 100, bottom: 100,
            left: 0, width: 100, height: 100, toJSON: (): object => ({})
        });
        svg.getScreenCTM = (): DOMMatrix => new DOMMatrix().translate(0, -scrollTop);

        const makeSystem: (key: string, top: number) => SVGGElement = (key: string, top: number): SVGGElement => {
            const group: SVGGElement = document.createElementNS("http://www.w3.org/2000/svg", "g");
            group.classList.add("osmd-system");
            group.dataset.osmdSystemKey = key;
            group.dataset.osmdSystemTop = top.toString();
            group.dataset.osmdSystemBottom = (top + 50).toString();
            svg.appendChild(group);
            return group;
        };
        const first: SVGGElement = makeSystem("1:0", 0);
        const distant: SVGGElement = makeSystem("1:1", 500);
        distant.style.opacity = "0.42";
        let clicks: number = 0;
        distant.addEventListener("click", (): number => ++clicks);

        const controller: SystemVirtualizationController = new SystemVirtualizationController(container);
        controller.enable({ scrollElement, overscanViewports: 0 });
        expect(first.isConnected).to.equal(true);
        expect(distant.isConnected).to.equal(false);
        expect(controller.stats).to.deep.include({
            totalSystems: 2,
            materializedSystems: 2,
            attachedSystems: 1,
            detachedSystems: 1,
            unmaterializedSystems: 0
        });

        scrollTop = 450;
        controller.updateNow();
        expect(first.isConnected).to.equal(false);
        expect(distant.isConnected).to.equal(true);
        expect(distant.style.opacity).to.equal("0.42");
        distant.dispatchEvent(new MouseEvent("click"));
        expect(clicks).to.equal(1);

        controller.disable();
        expect(first.isConnected).to.equal(true);
        expect(distant.isConnected).to.equal(true);
        scrollElement.remove();
    });

    it("evicts distant SVG groups and redraws a visited system", () => {
        const container: HTMLDivElement = document.createElement("div");
        const scrollElement: HTMLDivElement = document.createElement("div");
        const svg: SVGSVGElement = document.createElementNS("http://www.w3.org/2000/svg", "svg");
        document.body.appendChild(scrollElement);
        scrollElement.appendChild(container);
        container.appendChild(svg);
        let scrollTop: number = 0;
        scrollElement.getBoundingClientRect = (): DOMRect => ({
            x: 0, y: 0, top: 0, right: 100, bottom: 100,
            left: 0, width: 100, height: 100, toJSON: (): object => ({})
        });
        svg.getScreenCTM = (): DOMMatrix => new DOMMatrix().translate(0, -scrollTop);
        const created: Map<string, SVGGElement[]> = new Map();
        const evicted: string[] = [];
        const makeSystem: (key: string) => SVGGElement = (key: string): SVGGElement => {
            const top: number = Number.parseInt(key.split(":")[1], 10) * 200;
            const group: SVGGElement = document.createElementNS("http://www.w3.org/2000/svg", "g");
            group.classList.add("osmd-system");
            group.dataset.osmdSystemKey = key;
            group.dataset.osmdSystemTop = top.toString();
            group.dataset.osmdSystemBottom = (top + 50).toString();
            group.appendChild(document.createElementNS("http://www.w3.org/2000/svg", "path"));
            svg.appendChild(group);
            created.set(key, [...(created.get(key) ?? []), group]);
            return group;
        };
        const controller: SystemVirtualizationController = new SystemVirtualizationController(container);
        controller.enable({ scrollElement, overscanViewports: 0, maxCachedSvgNodes: 2 });
        controller.configureExpectedSystems(
            [0, 1, 2].map(index => ({ key: `1:${index}`, svg, top: index * 200, bottom: index * 200 + 50 })),
            keys => keys.map(makeSystem),
            key => evicted.push(key)
        );
        expect(controller.stats.materializedSystems).to.equal(1);
        scrollTop = 200;
        controller.updateNow();
        expect(evicted).to.deep.equal(["1:0"]);
        expect(controller.stats).to.deep.include({ materializedSystems: 1, retainedSvgNodes: 2, evictedSystems: 1 });
        expect(created.get("1:0")[0].childElementCount).to.equal(0);
        scrollTop = 0;
        controller.updateNow();
        expect(created.get("1:0")).to.have.length(2);
        expect(created.get("1:0")[1]).to.not.equal(created.get("1:0")[0]);
        expect(controller.stats.retainedSvgNodes).to.equal(2);
        controller.disable();
        scrollElement.remove();
    });
});
