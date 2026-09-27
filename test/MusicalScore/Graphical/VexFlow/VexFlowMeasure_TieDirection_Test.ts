/* eslint-disable @typescript-eslint/no-unused-expressions */
import {expect} from "chai";
import {IXmlElement} from "../../../../src/Common/FileIO/Xml";
import {GraphicalMusicSheet} from "../../../../src/MusicalScore/Graphical/GraphicalMusicSheet";
import {MusicSheet} from "../../../../src/MusicalScore/MusicSheet";
import {VexFlowMeasure} from "../../../../src/MusicalScore/Graphical/VexFlow/VexFlowMeasure";
import {VexFlowMusicSheetCalculator} from "../../../../src/MusicalScore/Graphical/VexFlow/VexFlowMusicSheetCalculator";
import {MusicSheetReader} from "../../../../src/MusicalScore/ScoreIO/MusicSheetReader";
import {TestUtils} from "../../../Util/TestUtils";

describe("VexFlow Measure - Automatic tie direction", () => {
   it("keeps clear ties shallow and flips the lower tie away from notes", (done: Mocha.Done) => {
      const filename: string = "test_tie_voice_order_k310.musicxml";
      const score: Document = TestUtils.getScore(filename);
      expect(score).to.not.be.undefined;
      const reader: MusicSheetReader = new MusicSheetReader();
      const calc: VexFlowMusicSheetCalculator = new VexFlowMusicSheetCalculator(reader.rules);
      const sheet: MusicSheet = reader.createMusicSheet(new IXmlElement(TestUtils.getPartWiseElement(score)), filename);
      const graphicalSheet: GraphicalMusicSheet = new GraphicalMusicSheet(sheet, calc);
      calc.calculate();

      const firstMeasure: VexFlowMeasure = graphicalSheet.MeasureList[0][0] as VexFlowMeasure;
      const secondMeasure: VexFlowMeasure = graphicalSheet.MeasureList[1][0] as VexFlowMeasure;
      expect(firstMeasure.vfTies.length).to.equal(1);
      expect(secondMeasure.vfTies.length).to.be.at.least(2);

      // The high C5 tie follows the clearer upper route.
      const upperTie: any = firstMeasure.vfTies[0];
      upperTie.direction = 1;
      (firstMeasure as any).placeTieAroundNotes(upperTie);
      expect(upperTie.direction).to.equal(-1);
      expect(upperTie.render_options.cp1).to.equal(8);

      // The A4 tie already belongs below the upper melody and stays there.
      const lowerTie: any = secondMeasure.vfTies[0];
      lowerTie.direction = 1;
      (secondMeasure as any).placeTieAroundNotes(lowerTie);
      expect(lowerTie.direction).to.equal(1);
      expect(lowerTie.render_options.cp1).to.equal(8);
      done();
   });

   it("preserves a tie direction specified by MusicXML", (done: Mocha.Done) => {
      const filename: string = "test_tie_voice_order_k310.musicxml";
      const score: Document = TestUtils.getScore(filename).cloneNode(true) as Document;
      const firstStart: Element = Array.from(score.getElementsByTagName("tied"))
         .find((element: Element): boolean => element.getAttribute("type") === "start");
      firstStart.setAttribute("placement", "below");
      const reader: MusicSheetReader = new MusicSheetReader();
      const calc: VexFlowMusicSheetCalculator = new VexFlowMusicSheetCalculator(reader.rules);
      const sheet: MusicSheet = reader.createMusicSheet(new IXmlElement(TestUtils.getPartWiseElement(score)), filename);
      const graphicalSheet: GraphicalMusicSheet = new GraphicalMusicSheet(sheet, calc);
      calc.calculate();
      const measure: VexFlowMeasure = graphicalSheet.MeasureList[0][0] as VexFlowMeasure;
      const tie: any = measure.vfTies[0];
      expect(tie.direction).to.equal(1);
      expect(tie.directionFromXml).to.equal(true);
      (measure as any).placeTieAroundNotes(tie);
      expect(tie.direction).to.equal(1);
      done();
   });
});
