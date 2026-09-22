import { expect } from "chai";
import { OpenSheetMusicDisplay } from "../../src/OpenSheetMusicDisplay/OpenSheetMusicDisplay";
import { VexFlowGraphicalNote } from "../../src/MusicalScore/Graphical/VexFlow/VexFlowGraphicalNote";
import { generateLargePianoScore } from "./LargeScoreFixtures";
import { TestUtils } from "../Util/TestUtils";

describe("Large score fixtures", () => {
    it("generate the requested measures with grace notes, ornaments, ties and meter changes", async () => {
        const container: HTMLElement = TestUtils.getDivElement(document);
        const osmd: OpenSheetMusicDisplay = TestUtils.createOpenSheetMusicDisplay(container);
        await osmd.load(generateLargePianoScore({
            measures: 40, graceEvery: 8, ornamentEvery: 6, tieEvery: 4, fingeringEvery: 4, meterChangeEvery: 10, systemBreakEvery: 4
        }));
        osmd.render();
        expect(osmd.Sheet.SourceMeasures.length).to.equal(40);
        const durations: number[] = osmd.Sheet.SourceMeasures.map(measure => measure.Duration.RealValue);
        expect(new Set(durations).size).to.equal(2);
        const notes: VexFlowGraphicalNote[] = osmd.GraphicSheet.MeasureList
            .flatMap(staves => staves)
            .flatMap(measure => measure.staffEntries)
            .flatMap(staffEntry => staffEntry.graphicalVoiceEntries)
            .flatMap(voiceEntry => voiceEntry.notes) as VexFlowGraphicalNote[];
        expect(notes.filter(note => note.isGraceNote()).length).to.equal(10);
        expect(notes.filter(note => note.hasOrnaments()).length).to.be.greaterThan(0);
        expect(notes.filter(note => note.sourceNote.NoteTie).length).to.be.greaterThan(0);
        expect(container.querySelectorAll(".pt-gracenote").length).to.be.greaterThan(0);
        expect(container.querySelectorAll(".pt-ornament").length).to.be.greaterThan(0);
        container.remove();
    });
});
