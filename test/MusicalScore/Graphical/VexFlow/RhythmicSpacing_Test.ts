import { expect } from "chai";
import { OpenSheetMusicDisplay } from "../../../../src/OpenSheetMusicDisplay/OpenSheetMusicDisplay";
import { VexFlowMeasure } from "../../../../src/MusicalScore/Graphical/VexFlow/VexFlowMeasure";
import { VexFlowVoiceEntry } from "../../../../src/MusicalScore/Graphical/VexFlow/VexFlowVoiceEntry";
import { TestUtils } from "../../../Util/TestUtils";
import Vex from "vexflow";

describe("Rhythmic spacing", () => {
   let container: HTMLDivElement;
   const note: (duration: number, type: string, staff: number, dot?: boolean) => string =
      (duration, type, staff, dot = false) => `<note><pitch><step>C</step><octave>${staff === 1 ? 5 : 3}</octave></pitch>
         <duration>${duration}</duration><voice>${staff}</voice><type>${type}</type>${dot ? "<dot/>" : ""}
         <stem>up</stem><staff>${staff}</staff></note>`;
   const xml: string = `<?xml version="1.0" encoding="UTF-8"?><score-partwise version="3.1">
      <part-list><score-part id="P1"><part-name>Piano</part-name></score-part></part-list><part id="P1"><measure number="1">
      <attributes><divisions>8</divisions><time><beats>4</beats><beat-type>4</beat-type></time><staves>2</staves>
         <clef number="1"><sign>G</sign><line>2</line></clef><clef number="2"><sign>F</sign><line>4</line></clef></attributes>
      ${note(4, "eighth", 1)}${note(8, "quarter", 1)}${note(4, "eighth", 1)}${note(4, "eighth", 1)}
      ${note(8, "quarter", 1)}${note(4, "eighth", 1)}<backup><duration>32</duration></backup>
      ${note(16, "half", 2)}${note(12, "quarter", 2, true)}${note(4, "eighth", 2)}
      </measure></part></score-partwise>`;

   beforeEach(() => {
      container = document.createElement("div");
      container.style.width = "600px";
      document.body.appendChild(container);
   });

   afterEach(() => container.remove());

   async function render(ratio?: number, score: string | Document = xml): Promise<OpenSheetMusicDisplay> {
      const osmd: OpenSheetMusicDisplay = new OpenSheetMusicDisplay(container, {
         backend: "svg", autoResize: false, drawTitle: false, drawComposer: false, drawPartNames: false,
         rhythmicSpacingRatio: ratio, stretchLastSystemLine: true
      });
      await osmd.load(score);
      osmd.render();
      return osmd;
   }

   function positions(osmd: OpenSheetMusicDisplay): number[][][] {
      return osmd.GraphicSheet.MeasureList.map(measures => measures.map(measure =>
         measure.staffEntries.map(entry => entry.PositionAndShape.RelativePosition.x)));
   }

   it("Preserves the default layout when disabled or given an invalid ratio", async () => {
      const original: number[][][] = positions(await render());
      for (const ratio of [0, -1, 0.5, 3.1, 4, NaN, Infinity]) {
         expect(positions(await render(ratio))).to.deep.equal(original);
      }
   });

   it("Gives quarter intervals 1.5 times the space of eighth intervals across screen widths", async () => {
      for (const width of [390, 600, 1000]) {
         container.style.width = `${width}px`;
         const osmd: OpenSheetMusicDisplay = await render(1.5);
         const x: number[] = (osmd.GraphicSheet.MeasureList[0][0] as VexFlowMeasure).vfVoices[1]
            .getTickables().map(tickable => (tickable as Vex.Flow.Note).getTickContext().getX());
         expect((x[2] - x[1]) / (x[1] - x[0])).to.be.closeTo(1.5, 0.00001);
         for (const measure of osmd.GraphicSheet.MeasureList[0]) {
            const voices: VexFlowMeasure["vfVoices"] = (measure as VexFlowMeasure).vfVoices;
            for (const voice of Object.values(voices)) {
               const notes: Vex.Flow.Tickable[] = voice.getTickables();
               for (let i: number = 1; i < notes.length; i++) {
                  const left: Vex.Flow.Note = notes[i - 1] as Vex.Flow.Note;
                  const right: Vex.Flow.Note = notes[i] as Vex.Flow.Note;
                  expect(right.getAbsoluteX()).to.be.greaterThan(left.getAbsoluteX());
                  const leftBox: Vex.Flow.BoundingBox = left.getBoundingBox();
                  const rightBox: Vex.Flow.BoundingBox = right.getBoundingBox();
                  expect(leftBox.getX() + leftBox.getW()).to.be.at.most(rightBox.getX() + 0.00001);
               }
            }
         }
      }
   });

   it("Keeps simultaneous onsets aligned between staves and can restore the original layout", async () => {
      const original: number[][][] = positions(await render());
      const osmd: OpenSheetMusicDisplay = await render(1.5);
      const measures: VexFlowMeasure[] = osmd.GraphicSheet.MeasureList[0] as VexFlowMeasure[];
      for (const lower of measures[1].staffEntries) {
         const upper: typeof lower = measures[0].staffEntries.find(entry =>
            entry.sourceStaffEntry.Timestamp.Equals(lower.sourceStaffEntry.Timestamp));
         const upperNote: Vex.Flow.StemmableNote = (upper.graphicalVoiceEntries[0] as VexFlowVoiceEntry).vfStaveNote;
         const lowerNote: Vex.Flow.StemmableNote = (lower.graphicalVoiceEntries[0] as VexFlowVoiceEntry).vfStaveNote;
         expect(upperNote.getTickContext().getX()).to.equal(lowerNote.getTickContext().getX());
      }
      const spaced: number[][][] = positions(osmd);
      osmd.render();
      expect(positions(osmd)).to.deep.equal(spaced);
      osmd.setOptions({ rhythmicSpacingRatio: 0 });
      osmd.render();
      expect(positions(osmd)).to.deep.equal(original);
   });

   it("Accepts ratios up to three and assigns more space to longer intervals", async () => {
      for (const width of [390, 1000]) {
         container.style.width = `${width}px`;
         const base: OpenSheetMusicDisplay = await render(2);
         const baseX: number[] = (base.GraphicSheet.MeasureList[0][0] as VexFlowMeasure).vfVoices[1]
            .getTickables().map(tickable => (tickable as Vex.Flow.Note).getTickContext().getX());
         for (const ratio of [2.5, 3]) {
            const osmd: OpenSheetMusicDisplay = await render(ratio);
            const x: number[] = (osmd.GraphicSheet.MeasureList[0][0] as VexFlowMeasure).vfVoices[1]
               .getTickables().map(tickable => (tickable as Vex.Flow.Note).getTickContext().getX());
            expect(x[2] - x[1]).to.be.at.least(baseX[2] - baseX[1] - 0.00001);
            expect(x[1] - x[0]).to.be.at.most(baseX[1] - baseX[0]);
            expect(x.every(Number.isFinite)).to.equal(true);
            const notes: Vex.Flow.Tickable[] = (osmd.GraphicSheet.MeasureList[0][0] as VexFlowMeasure).vfVoices[1].getTickables();
            for (let index: number = 1; index < notes.length; index++) {
               const left: Vex.Flow.BoundingBox = (notes[index - 1] as Vex.Flow.Note).getBoundingBox();
               const right: Vex.Flow.BoundingBox = (notes[index] as Vex.Flow.Note).getBoundingBox();
               expect(left.getX() + left.getW()).to.be.at.most(right.getX() + 0.00001);
            }
            if (width === 1000) {
               expect(x[2] - x[1]).to.be.greaterThan(baseX[2] - baseX[1]);
               expect((x[2] - x[1]) / (x[1] - x[0])).to.be.closeTo(ratio, 0.00001);
            }
         }
      }
   });

   it("Keeps complex scores stable through a second render", async () => {
      for (const ratio of [1.5, 3]) {
         for (const filename of ["test_tuplet_multivoice_alignment.musicxml", "test_grace_notes_after_main_note_1706.musicxml",
            "test_lyrics_overlap_pickup_anacrusis_dash_minden_m1.musicxml", "test_voice_gaps_of_a_whole_note_or_more.musicxml"]) {
            const osmd: OpenSheetMusicDisplay = await render(ratio, TestUtils.getScore(filename));
            const first: number[][][] = positions(osmd);
            expect(first.flat(2).every(Number.isFinite), filename).to.equal(true);
            osmd.render();
            expect(positions(osmd), filename).to.deep.equal(first);
         }
      }
   });
});
