import { expect } from "chai";
import { OpenSheetMusicDisplay } from "../../src/OpenSheetMusicDisplay/OpenSheetMusicDisplay";
import { GraphicalNote } from "../../src/MusicalScore/Graphical/GraphicalNote";
import { TestUtils } from "../Util/TestUtils";

function notesInBeatRange(osmd: OpenSheetMusicDisplay, measureListIndex: number, fromBeat: number, beatCount: number): Set<GraphicalNote> {
    const start: number = Math.max(0, fromBeat) * 0.25 - 1e-6;
    const end: number = Math.max(0, fromBeat + beatCount) * 0.25 - 1e-6;
    const notes: Set<GraphicalNote> = new Set<GraphicalNote>();
    for (const container of osmd.GraphicSheet.VerticalGraphicalStaffEntryContainers) {
        for (const staffEntry of container?.StaffEntries ?? []) {
            if (staffEntry?.parentMeasure?.parentSourceMeasure?.measureListIndex !== measureListIndex) {
                continue;
            }
            const time: number = staffEntry.relInMeasureTimestamp?.RealValue ?? 0;
            if (time < start || time >= end) {
                continue;
            }
            for (const voiceEntry of staffEntry.graphicalVoiceEntries) {
                for (const note of voiceEntry.notes) {
                    notes.add(note);
                }
            }
        }
    }
    return notes;
}

function allNotes(osmd: OpenSheetMusicDisplay): GraphicalNote[] {
    return osmd.GraphicSheet.MeasureList
        .flatMap(staves => staves)
        .filter(measure => !!measure)
        .flatMap(measure => measure.staffEntries)
        .flatMap(staffEntry => staffEntry.graphicalVoiceEntries)
        .flatMap(voiceEntry => voiceEntry.notes);
}

describe("Read-ahead staff entry index", () => {
    let container: HTMLElement;
    let osmd: OpenSheetMusicDisplay;

    beforeEach(async () => {
        container = TestUtils.getDivElement(document);
        osmd = TestUtils.createOpenSheetMusicDisplay(container);
        await osmd.load(TestUtils.getScore("MuzioClementi_SonatinaOpus36No1_Part1.xml"));
        osmd.render();
    });

    afterEach(() => {
        container.remove();
    });

    it("hides exactly the notes a full scan selects for every beat range", () => {
        const measureCount: number = osmd.Sheet.SourceMeasures.length;
        for (let measureListIndex: number = 0; measureListIndex < measureCount; measureListIndex++) {
            for (const [fromBeat, beatCount] of [[0, 1], [0.5, 1], [1, 2], [2.75, 0.5], [0, 8], [3.99, 1]]) {
                osmd.resetReadAheadOpacity();
                osmd.hideReadAheadMeasureBeatRange(measureListIndex, fromBeat, beatCount, 0);
                const expected: Set<GraphicalNote> = notesInBeatRange(osmd, measureListIndex, fromBeat, beatCount);
                for (const note of allNotes(osmd)) {
                    const hidden: boolean = note.opacity === 0;
                    expect(hidden, `measure ${measureListIndex} range ${fromBeat}+${beatCount}`).to.equal(expected.has(note));
                }
            }
        }
        osmd.resetReadAheadOpacity();
    });

    it("rebuilds after the graphical score is recalculated", () => {
        const firstGeneration: number = osmd.LayoutGeneration;
        expect(firstGeneration).to.be.greaterThan(0);
        osmd.hideReadAheadMeasureBeatRange(1, 0, 4, 0);
        osmd.resetReadAheadOpacity();

        osmd.Zoom = 0.8;
        osmd.render();
        expect(osmd.LayoutGeneration).to.not.equal(firstGeneration);

        osmd.hideReadAheadMeasureBeatRange(1, 0, 4, 0);
        const expected: Set<GraphicalNote> = notesInBeatRange(osmd, 1, 0, 4);
        expect(expected.size).to.be.greaterThan(0);
        for (const note of expected) {
            expect(note.opacity).to.equal(0);
        }
        osmd.resetReadAheadOpacity();
    });

    it("never shares a generation between loaded scores", async () => {
        const previous: number = osmd.LayoutGeneration;
        await osmd.load(TestUtils.getScore("MuzioClementi_SonatinaOpus36No1_Part2.xml"));
        expect(osmd.LayoutGeneration).to.not.equal(previous);
        osmd.render();
        osmd.hideReadAheadMeasureBeatRange(0, 0, 4, 0);
        for (const note of notesInBeatRange(osmd, 0, 0, 4)) {
            expect(note.opacity).to.equal(0);
        }
    });
});
