import { expect } from "chai";
import { OpenSheetMusicDisplay } from "../../src/OpenSheetMusicDisplay/OpenSheetMusicDisplay";
import { VexFlowGraphicalNote } from "../../src/MusicalScore/Graphical/VexFlow/VexFlowGraphicalNote";
import { GraphicalNote } from "../../src/MusicalScore/Graphical/GraphicalNote";
import { TestUtils } from "../Util/TestUtils";

const DIMMED_ROOT_CLASS: string = "pt-practice-notation-dimmed";
const DIMMED_OPACITY: number = 0.15;

function allGraphicalNotes(osmd: OpenSheetMusicDisplay): VexFlowGraphicalNote[] {
    const notes: VexFlowGraphicalNote[] = [];
    for (const page of osmd.GraphicSheet.MusicPages) {
        for (const system of page.MusicSystems) {
            for (const staffLine of system.StaffLines) {
                for (const measure of staffLine.Measures) {
                    for (const staffEntry of measure?.staffEntries ?? []) {
                        for (const voiceEntry of staffEntry.graphicalVoiceEntries) {
                            for (const note of voiceEntry.notes) {
                                notes.push(note as VexFlowGraphicalNote);
                            }
                        }
                    }
                }
            }
        }
    }
    return notes;
}

function effectiveOpacity(element: Element, root: Element): number {
    let opacity: number = 1;
    let current: Element | null = element;
    while (current && current !== root.parentElement) {
        opacity *= Number.parseFloat(getComputedStyle(current).opacity || "1");
        current = current.parentElement;
    }
    return opacity;
}

function expectGraceGroupsDimmed(container: HTMLElement): number {
    const groups: Element[] = Array.from(container.querySelectorAll(".vf-gracenotegroup"));
    for (const group of groups) {
        for (const path of Array.from(group.querySelectorAll("path"))) {
            expect(effectiveOpacity(path, container), `grace group ${path.parentElement?.getAttribute("class")} path`)
                .to.be.closeTo(DIMMED_OPACITY, 0.01);
        }
    }
    return groups.length;
}

interface MountedScore {
    scrollElement: HTMLDivElement;
    container: HTMLDivElement;
    style: HTMLStyleElement;
}

const DIMMING_RULE: string = [".pt-gracenote", ".pt-ornament"]
    .map(selector => `.${DIMMED_ROOT_CLASS} ${selector}:not([opacity="0"])`)
    .join(", ");

function mountScrollableScore(width: number, height: number): MountedScore {
    const style: HTMLStyleElement = document.createElement("style");
    style.textContent = `${DIMMING_RULE} { opacity: ${DIMMED_OPACITY}; }`;
    document.head.appendChild(style);
    const scrollElement: HTMLDivElement = document.createElement("div");
    scrollElement.style.width = `${width}px`;
    scrollElement.style.height = `${height}px`;
    scrollElement.style.overflow = "auto";
    const container: HTMLDivElement = document.createElement("div");
    container.style.width = `${width}px`;
    scrollElement.appendChild(container);
    document.body.appendChild(scrollElement);
    return { scrollElement, container, style };
}

function expectGraceNotesDimmed(notes: GraphicalNote[], container: HTMLElement): number {
    let checked: number = 0;
    for (const note of notes as VexFlowGraphicalNote[]) {
        if (!note.isGraceNote()) {
            continue;
        }
        const element: SVGGElement = note.getSVGGElement();
        if (!element?.isConnected) {
            continue;
        }
        const parts: Element[] = [
            ...note.getNoteheadSVGs(),
            note.getStemSVG(),
            note.getFlagSVG(),
            ...note.getBeamSVGs(),
            ...note.getLedgerLineSVGs()
        ].filter(part => !!part);
        expect(parts.length).to.be.greaterThan(0);
        for (const part of parts) {
            expect(effectiveOpacity(part, container), `grace ${note.getSVGId()} ${part.getAttribute("class")} opacity`)
                .to.be.closeTo(DIMMED_OPACITY, 0.01);
        }
        checked++;
    }
    return checked;
}

describe("Grace note and ornament practice classes", () => {
    it("classifies grace notes in a full render", async () => {
        const { scrollElement, container, style } = mountScrollableScore(800, 600);
        const osmd: OpenSheetMusicDisplay = new OpenSheetMusicDisplay(container, { autoResize: false });
        await osmd.load(TestUtils.getScore("OSMD_function_test_GraceNotes.xml"));
        osmd.render();
        container.classList.add(DIMMED_ROOT_CLASS);
        const graceCount: number = allGraphicalNotes(osmd).filter(note => note.isGraceNote()).length;
        expect(graceCount).to.be.greaterThan(0);
        expect(expectGraceNotesDimmed(allGraphicalNotes(osmd), container)).to.equal(graceCount);
        expect(expectGraceGroupsDimmed(container)).to.be.greaterThan(0);
        expect(container.querySelector(".vf-gracenotegroup .vf-stavetie, .vf-gracenotegroup .vf-beam")).to.not.equal(null);
        container.classList.remove(DIMMED_ROOT_CLASS);
        for (const note of allGraphicalNotes(osmd).filter(n => n.isGraceNote())) {
            for (const notehead of note.getNoteheadSVGs()) {
                expect(effectiveOpacity(notehead, container)).to.equal(1);
            }
        }
        scrollElement.remove();
        style.remove();
    });

    it("classifies grace notes in systems materialized later by virtualization", async () => {
        const { scrollElement, container, style } = mountScrollableScore(360, 200);
        const osmd: OpenSheetMusicDisplay = new OpenSheetMusicDisplay(container, { autoResize: false });
        osmd.enableSystemVirtualization({ scrollElement, overscanViewports: 0 });
        await osmd.load(TestUtils.getScore("OSMD_function_test_GraceNotes.xml"));
        container.classList.add(DIMMED_ROOT_CLASS);
        osmd.renderVirtualized({ initialSystems: 1 });
        const totalSystems: number = osmd.GraphicSheet.MusicPages
            .reduce((sum, page): number => sum + page.MusicSystems.length, 0);
        expect(totalSystems).to.be.greaterThan(1);
        const graceCount: number = allGraphicalNotes(osmd).filter(note => note.isGraceNote()).length;
        let checked: number = expectGraceNotesDimmed(allGraphicalNotes(osmd), container);
        expect(checked).to.be.lessThan(graceCount);
        for (let top: number = 0; top <= scrollElement.scrollHeight; top += 100) {
            scrollElement.scrollTop = top;
            osmd.updateSystemVirtualization();
        }
        osmd.disableSystemVirtualization();
        checked = expectGraceNotesDimmed(allGraphicalNotes(osmd), container);
        expect(checked).to.equal(graceCount);
        expectGraceGroupsDimmed(container);
        scrollElement.remove();
        style.remove();
    });

    it("classifies ornaments without dimming the main notehead", async () => {
        const { scrollElement, container, style } = mountScrollableScore(800, 600);
        const osmd: OpenSheetMusicDisplay = new OpenSheetMusicDisplay(container, { autoResize: false });
        await osmd.load(TestUtils.getScore("OSMD_function_test_Ornaments.xml"));
        osmd.render();
        container.classList.add(DIMMED_ROOT_CLASS);
        const ornamented: VexFlowGraphicalNote[] = allGraphicalNotes(osmd).filter(note => note.hasOrnaments());
        expect(ornamented.length).to.be.greaterThan(0);
        for (const note of ornamented) {
            const element: SVGGElement = note.getSVGGElement();
            const ornamentNodes: Element[] = Array.from(element.querySelectorAll(".pt-ornament"));
            expect(ornamentNodes.length, `ornament ${note.getSVGId()}`).to.be.greaterThan(0);
            for (const ornamentNode of ornamentNodes) {
                expect(ornamentNode.classList.contains("vf-ornament")).to.equal(true);
                expect(effectiveOpacity(ornamentNode, container)).to.be.closeTo(DIMMED_OPACITY, 0.01);
            }
            for (const notehead of note.getNoteheadSVGs()) {
                expect(effectiveOpacity(notehead, container)).to.equal(1);
            }
            for (const accidental of Array.from(element.querySelectorAll(".vf-modifiers > :not(.vf-ornament)"))) {
                expect(effectiveOpacity(accidental, container)).to.equal(1);
            }
        }

        const hidden: VexFlowGraphicalNote = ornamented[0];
        hidden.setOpacity(0);
        for (const ornamentNode of Array.from(hidden.getSVGGElement().querySelectorAll(".pt-ornament"))) {
            expect(effectiveOpacity(ornamentNode, container)).to.equal(0);
        }
        hidden.setOpacity(1);
        for (const ornamentNode of Array.from(hidden.getSVGGElement().querySelectorAll(".pt-ornament"))) {
            expect(effectiveOpacity(ornamentNode, container)).to.be.closeTo(DIMMED_OPACITY, 0.01);
        }
        scrollElement.remove();
        style.remove();
    });

    it("hides grace notes with their main note during read-ahead opacity changes", async () => {
        const { scrollElement, container, style } = mountScrollableScore(800, 600);
        const osmd: OpenSheetMusicDisplay = new OpenSheetMusicDisplay(container, { autoResize: false });
        await osmd.load(TestUtils.getScore("OSMD_function_test_GraceNotes.xml"));
        osmd.render();
        container.classList.add(DIMMED_ROOT_CLASS);
        const notes: VexFlowGraphicalNote[] = allGraphicalNotes(osmd);
        const mainNote: VexFlowGraphicalNote = notes.find(note =>
            !note.isGraceNote() && note.getSVGGElement()?.querySelector(".vf-gracenotegroup"));
        expect(mainNote).to.not.equal(undefined);
        const graceGroup: Element = mainNote.getSVGGElement().querySelector(".vf-gracenotegroup");
        for (const note of notes.filter(n => n.parentVoiceEntry.parentStaffEntry === mainNote.parentVoiceEntry.parentStaffEntry)) {
            note.setOpacity(0);
        }
        for (const path of Array.from(graceGroup.querySelectorAll("path"))) {
            expect(effectiveOpacity(path, container)).to.equal(0);
        }
        for (const note of notes.filter(n => n.parentVoiceEntry.parentStaffEntry === mainNote.parentVoiceEntry.parentStaffEntry)) {
            note.setOpacity(1);
        }
        for (const path of Array.from(graceGroup.querySelectorAll("path"))) {
            expect(effectiveOpacity(path, container)).to.be.closeTo(DIMMED_OPACITY, 0.01);
        }
        scrollElement.remove();
        style.remove();
    });

    for (const fixture of [
        "test_grace_note_arpeggio_ysaye.musicxml",
        "test_grace_note_fingerings_position.musicxml",
        "test_grace_slash.musicxml",
        "test_graceslash_simple.musicxml",
        "test_octaveshift_multiline_grace_notes.musicxml",
        "test_octaveshift_stop_after_grace_notes.musicxml",
        "OSMD_function_test_Ornaments.xml"
    ]) {
        it(`dims every grace note part in ${fixture}`, async () => {
            const { scrollElement, container, style } = mountScrollableScore(800, 600);
            const osmd: OpenSheetMusicDisplay = new OpenSheetMusicDisplay(container, { autoResize: false });
            await osmd.load(TestUtils.getScore(fixture));
            osmd.render();
            container.classList.add(DIMMED_ROOT_CLASS);
            const graceCount: number = allGraphicalNotes(osmd).filter(note => note.isGraceNote()).length;
            expect(expectGraceNotesDimmed(allGraphicalNotes(osmd), container)).to.equal(graceCount);
            expectGraceGroupsDimmed(container);
            for (const note of allGraphicalNotes(osmd).filter(n => !n.isGraceNote())) {
                for (const notehead of note.getNoteheadSVGs()) {
                    expect(effectiveOpacity(notehead, container), `main ${note.getSVGId()}`).to.equal(1);
                }
            }
            scrollElement.remove();
            style.remove();
        });
    }
});
