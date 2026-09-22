import { expect } from "chai";
import { OpenSheetMusicDisplay } from "../../src/OpenSheetMusicDisplay/OpenSheetMusicDisplay";
import { ISystemLifecycleEvent } from "../../src/OpenSheetMusicDisplay/SystemVirtualizationController";
import { generateLargePianoScore } from "../performance/LargeScoreFixtures";
import { VexFlowGraphicalNote } from "../../src/MusicalScore/Graphical/VexFlow/VexFlowGraphicalNote";

function expectedAttachedKeys(osmd: OpenSheetMusicDisplay, scrollElement: HTMLElement, overscan: number): Set<string> {
    const rect: DOMRect = scrollElement.getBoundingClientRect();
    const minY: number = rect.top - rect.height * overscan;
    const maxY: number = rect.bottom + rect.height * overscan;
    const keys: Set<string> = new Set<string>();
    for (let pageIndex: number = 0; pageIndex < osmd.GraphicSheet.MusicPages.length; pageIndex++) {
        const page: any = osmd.GraphicSheet.MusicPages[pageIndex];
        const svg: SVGSVGElement = osmd.Drawer.Backends[pageIndex].getRenderElement().querySelector("svg");
        const matrix: DOMMatrix = svg.getScreenCTM();
        for (let systemIndex: number = 0; systemIndex < page.MusicSystems.length; systemIndex++) {
            const bounds: any = osmd.Drawer.getSystemPixelBounds(page.MusicSystems[systemIndex]);
            const top: number = new DOMPoint(0, bounds.y).matrixTransform(matrix).y;
            const bottom: number = new DOMPoint(0, bounds.y + bounds.height).matrixTransform(matrix).y;
            if (bottom >= minY && top <= maxY) {
                keys.add(`${page.PageNumber}:${systemIndex}`);
            }
        }
    }
    return keys;
}

function connectedKeys(container: HTMLElement): Set<string> {
    return new Set(Array.from(container.querySelectorAll<SVGGElement>("g.osmd-system"))
        .map(group => group.dataset.osmdSystemKey));
}

describe("System virtualization lookup", () => {
    for (const overscan of [0, 1]) {
        it(`attaches exactly the systems in the window and reports every change (overscan ${overscan})`, async () => {
            const scrollElement: HTMLDivElement = document.createElement("div");
            scrollElement.style.cssText = "width:900px;height:500px;overflow:auto";
            const container: HTMLDivElement = document.createElement("div");
            container.style.width = "900px";
            scrollElement.appendChild(container);
            document.body.appendChild(scrollElement);
            const osmd: OpenSheetMusicDisplay = new OpenSheetMusicDisplay(container, { autoResize: false });
            osmd.enableSystemVirtualization({ scrollElement, overscanViewports: overscan });
            const tracked: Set<string> = new Set<string>();
            const materialized: Set<string> = new Set<string>();
            osmd.addSystemLifecycleListener((event: ISystemLifecycleEvent): void => {
                expect(event.keys.length).to.equal(event.roots.length);
                for (const key of event.keys) {
                    if (event.type === "detached") {
                        tracked.delete(key);
                    } else {
                        tracked.add(key);
                    }
                    if (event.type === "materialized") {
                        expect(materialized.has(key), `${key} materialized twice`).to.equal(false);
                        materialized.add(key);
                    }
                }
            });
            await osmd.load(generateLargePianoScore({ measures: 120, graceEvery: 8, tieEvery: 4 }));
            osmd.renderVirtualized({ initialSystems: 2 });

            let seed: number = 5;
            const positions: number[] = [0, scrollElement.scrollHeight];
            for (let i: number = 0; i < 40; i++) {
                seed = (seed * 9301 + 49297) % 233280;
                positions.push(Math.floor((seed / 233280) * scrollElement.scrollHeight));
            }
            for (const top of positions) {
                scrollElement.scrollTop = top;
                osmd.updateSystemVirtualization();
                const expected: Set<string> = expectedAttachedKeys(osmd, scrollElement, overscan);
                const connected: Set<string> = connectedKeys(container);
                for (const key of expected) {
                    if (materialized.has(key)) {
                        expect(connected.has(key), `${key} attached at ${top}`).to.equal(true);
                    }
                }
                for (const key of connected) {
                    expect(expected.has(key), `${key} detached at ${top}`).to.equal(true);
                }
                expect(Array.from(tracked).sort()).to.deep.equal(Array.from(connected).sort());
                expect(osmd.SystemVirtualizationStats.attachedSystems).to.equal(connected.size);
            }
            osmd.disableSystemVirtualization();
            scrollElement.remove();
        });
    }

    it("keeps staff opacity overrides on systems that were detached when they changed", async () => {
        const scrollElement: HTMLDivElement = document.createElement("div");
        scrollElement.style.cssText = "width:900px;height:400px;overflow:auto";
        const container: HTMLDivElement = document.createElement("div");
        container.style.width = "900px";
        scrollElement.appendChild(container);
        document.body.appendChild(scrollElement);
        const osmd: OpenSheetMusicDisplay = new OpenSheetMusicDisplay(container, { autoResize: false });
        osmd.enableSystemVirtualization({ scrollElement, overscanViewports: 0 });
        await osmd.load(generateLargePianoScore({ measures: 60 }));
        osmd.renderVirtualized({ initialSystems: 1 });
        scrollElement.scrollTop = scrollElement.scrollHeight;
        osmd.updateSystemVirtualization();
        scrollElement.scrollTop = 0;
        osmd.updateSystemVirtualization();

        osmd.blurStaff(1, 0.3);
        scrollElement.scrollTop = scrollElement.scrollHeight;
        osmd.updateSystemVirtualization();
        const lowerStaves: SVGGElement[] = Array.from(container.querySelectorAll<SVGGElement>(".staffline[data-staff-index=\"1\"]"));
        expect(lowerStaves.length).to.be.greaterThan(0);
        for (const staffline of lowerStaves) {
            expect(staffline.style.opacity).to.equal("0.3");
        }
        osmd.restoreStaff(1);
        scrollElement.scrollTop = 0;
        osmd.updateSystemVirtualization();
        for (const staffline of Array.from(container.querySelectorAll<SVGGElement>(".staffline[data-staff-index=\"1\"]"))) {
            expect(staffline.style.opacity).to.equal("");
        }
        osmd.disableSystemVirtualization();
        scrollElement.remove();
    });

    it("writes state changed while a system was detached when it is reattached", async () => {
        const scrollElement: HTMLDivElement = document.createElement("div");
        scrollElement.style.cssText = "width:900px;height:400px;overflow:auto";
        const container: HTMLDivElement = document.createElement("div");
        container.style.width = "900px";
        scrollElement.appendChild(container);
        document.body.appendChild(scrollElement);
        const osmd: OpenSheetMusicDisplay = new OpenSheetMusicDisplay(container, { autoResize: false });
        osmd.enableSystemVirtualization({ scrollElement, overscanViewports: 0 });
        await osmd.load(generateLargePianoScore({ measures: 60 }));
        osmd.renderVirtualized({ initialSystems: 1 });
        const firstSystemNotes: VexFlowGraphicalNote[] = osmd.GraphicSheet.MusicPages[0].MusicSystems[0].StaffLines
            .flatMap(staffLine => staffLine.Measures)
            .flatMap(measure => measure.staffEntries)
            .flatMap(staffEntry => staffEntry.graphicalVoiceEntries)
            .flatMap(voiceEntry => voiceEntry.notes) as VexFlowGraphicalNote[];
        const note: VexFlowGraphicalNote = firstSystemNotes[0];
        const notehead: () => Element = (): Element => note.getNoteheadSVGs()[0].querySelector("path");
        note.setOpacity(0);
        expect(notehead().getAttribute("opacity")).to.equal("0");

        scrollElement.scrollTop = scrollElement.scrollHeight;
        osmd.updateSystemVirtualization();
        expect(note.getSVGGElement().isConnected).to.equal(false);
        note.setOpacity(1);
        note.setColor("#00aa00");
        expect(notehead().getAttribute("opacity")).to.equal("0");
        expect(notehead().getAttribute("fill")).to.not.equal("#00aa00");

        scrollElement.scrollTop = 0;
        osmd.updateSystemVirtualization();
        expect(note.getSVGGElement().isConnected).to.equal(true);
        expect(notehead().getAttribute("opacity")).to.equal("1");
        expect(notehead().getAttribute("fill")).to.equal("#00aa00");

        osmd.disableSystemVirtualization();
        scrollElement.remove();
    });

    it("writes immediately when the whole score is outside the document", async () => {
        const container: HTMLDivElement = document.createElement("div");
        container.style.width = "900px";
        const osmd: OpenSheetMusicDisplay = new OpenSheetMusicDisplay(container, { autoResize: false });
        await osmd.load(generateLargePianoScore({ measures: 8 }));
        osmd.render();
        const note: VexFlowGraphicalNote = osmd.GraphicSheet.MeasureList[0][0].staffEntries[0]
            .graphicalVoiceEntries[0].notes[0] as VexFlowGraphicalNote;
        note.setOpacity(0.5);
        expect(note.getNoteheadSVGs()[0].querySelector("path").getAttribute("opacity")).to.equal("0.5");
    });

    const nextFrame: () => Promise<void> = (): Promise<void> =>
        new Promise<void>(resolve => requestAnimationFrame((): void => resolve()));

    async function mountVirtualized(budgets: { active: number, idle: number }): Promise<{
        osmd: OpenSheetMusicDisplay; scrollElement: HTMLDivElement;
    }> {
        const scrollElement: HTMLDivElement = document.createElement("div");
        scrollElement.style.cssText = "width:900px;height:400px;overflow:auto";
        const container: HTMLDivElement = document.createElement("div");
        container.style.width = "900px";
        scrollElement.appendChild(container);
        document.body.appendChild(scrollElement);
        const osmd: OpenSheetMusicDisplay = new OpenSheetMusicDisplay(container, { autoResize: false });
        osmd.enableSystemVirtualization({
            scrollElement, overscanViewports: 2,
            activeMaterializationBudgetMs: budgets.active, idleMaterializationBudgetMs: budgets.idle
        });
        await osmd.load(generateLargePianoScore({ measures: 120 }));
        osmd.renderVirtualized({ initialSystems: 1 });
        return { osmd, scrollElement };
    }

    it("queues offscreen systems nearest first, favoring the scroll direction", async () => {
        const { osmd, scrollElement } = await mountVirtualized({ active: 0, idle: 0 });
        const controller: any = (osmd as any).systemVirtualization;
        scrollElement.scrollTop = scrollElement.scrollHeight / 3;
        osmd.updateSystemVirtualization();
        scrollElement.scrollTop += 200;
        osmd.updateSystemVirtualization();
        const queue: string[] = [...controller.pendingMaterializationKeys];
        expect(queue.length).to.be.greaterThan(1);
        const indexOf: (key: string) => number = (key: string): number => Number(key.split(":")[1]);
        const visible: number[] = Array.from(scrollElement.querySelectorAll<SVGGElement>("g.osmd-system"))
            .map(group => indexOf(group.dataset.osmdSystemKey));
        const lastVisible: number = Math.max(...visible);
        expect(indexOf(queue[0])).to.equal(lastVisible + 1);
        expect(osmd.SystemVirtualizationStats.pendingMaterializations).to.equal(queue.length);
        osmd.disableSystemVirtualization();
        scrollElement.remove();
    });

    it("draws one queued system per frame with no budget and drains the queue with a large one", async () => {
        const small: { osmd: OpenSheetMusicDisplay, scrollElement: HTMLDivElement } = await mountVirtualized({ active: 0, idle: 0 });
        small.osmd.updateSystemVirtualization();
        const pendingBefore: number = small.osmd.SystemVirtualizationStats.pendingMaterializations;
        expect(pendingBefore).to.be.greaterThan(1);
        await nextFrame();
        await nextFrame();
        expect(small.osmd.SystemVirtualizationStats.pendingMaterializations).to.be.at.least(pendingBefore - 2);
        expect(small.osmd.SystemVirtualizationStats.averageMaterializationMs).to.be.greaterThan(0);
        small.osmd.disableSystemVirtualization();
        small.scrollElement.remove();

        const large: { osmd: OpenSheetMusicDisplay, scrollElement: HTMLDivElement } = await mountVirtualized({ active: 10000, idle: 10000 });
        large.osmd.updateSystemVirtualization();
        expect(large.osmd.SystemVirtualizationStats.pendingMaterializations).to.be.greaterThan(1);
        await nextFrame();
        await nextFrame();
        expect(large.osmd.SystemVirtualizationStats.pendingMaterializations).to.equal(0);
        large.osmd.disableSystemVirtualization();
        large.scrollElement.remove();
    });
});
