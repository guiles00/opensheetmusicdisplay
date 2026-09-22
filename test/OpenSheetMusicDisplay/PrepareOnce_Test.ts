import { expect } from "chai";
import { OpenSheetMusicDisplay } from "../../src/OpenSheetMusicDisplay/OpenSheetMusicDisplay";
import { MusicSheetCalculator } from "../../src/MusicalScore/Graphical/MusicSheetCalculator";
import { TestUtils } from "../Util/TestUtils";

describe("Graphical sheet preparation", () => {
    let prepareCount: number;
    const originalPrepare: () => void = MusicSheetCalculator.prototype.prepareGraphicalMusicSheet;

    beforeEach(() => {
        prepareCount = 0;
        MusicSheetCalculator.prototype.prepareGraphicalMusicSheet = function (this: MusicSheetCalculator): void {
            prepareCount++;
            originalPrepare.call(this);
        };
    });

    afterEach(() => {
        MusicSheetCalculator.prototype.prepareGraphicalMusicSheet = originalPrepare;
    });

    it("defers preparation from load until the graphic sheet is first used", async () => {
        const container: HTMLElement = TestUtils.getDivElement(document);
        const osmd: OpenSheetMusicDisplay = TestUtils.createOpenSheetMusicDisplay(container);
        await osmd.load(TestUtils.getScore("MuzioClementi_SonatinaOpus36No1_Part1.xml"));
        expect(prepareCount).to.equal(0);
        expect(osmd.GraphicSheet.MeasureList.length).to.equal(osmd.Sheet.SourceMeasures.length);
        expect(prepareCount).to.equal(1);
        osmd.render();
        expect(prepareCount).to.equal(1);
        container.remove();
    });

    it("prepares once for load plus the first render", async () => {
        const container: HTMLElement = TestUtils.getDivElement(document);
        const osmd: OpenSheetMusicDisplay = TestUtils.createOpenSheetMusicDisplay(container);
        await osmd.load(TestUtils.getScore("MuzioClementi_SonatinaOpus36No1_Part1.xml"));
        osmd.render();
        osmd.render();
        expect(prepareCount).to.equal(1);
        container.remove();
    });

    it("prepares again when the drawing range changes", async () => {
        const container: HTMLElement = TestUtils.getDivElement(document);
        const osmd: OpenSheetMusicDisplay = TestUtils.createOpenSheetMusicDisplay(container);
        await osmd.load(TestUtils.getScore("MuzioClementi_SonatinaOpus36No1_Part1.xml"));
        osmd.render();
        osmd.setOptions({ drawFromMeasureNumber: 3, drawUpToMeasureNumber: 6 });
        osmd.render();
        expect(prepareCount).to.equal(2);
        expect(osmd.GraphicSheet.MeasureList.length).to.be.greaterThan(0);
        osmd.setOptions({ drawFromMeasureNumber: 3, drawUpToMeasureNumber: 6 });
        osmd.render();
        expect(prepareCount).to.equal(2);
        container.remove();
    });

    for (const name of ["test_fermata_inverted_placement.musicxml", "test_lyrics_overlap_pickup_anacrusis_dash_minden_m1.musicxml"]) {
        it(`renders ${name} identically however often it was prepared`, async () => {
            const renderAfter: (preparations: number) => Promise<string> = async (preparations: number): Promise<string> => {
                const container: HTMLElement = TestUtils.getDivElement(document);
                container.style.width = "1000px";
                const osmd: OpenSheetMusicDisplay = TestUtils.createOpenSheetMusicDisplay(container);
                await osmd.load(TestUtils.getScore(name));
                for (let i: number = 0; i < preparations; i++) {
                    osmd.GraphicSheet.GetCalculator.prepareGraphicalMusicSheet();
                }
                osmd.render();
                const svg: string = container.innerHTML.replace(/auto\d+/g, "auto");
                container.remove();
                return svg;
            };
            const once: string = await renderAfter(0);
            expect(await renderAfter(1)).to.equal(once);
            expect(await renderAfter(2)).to.equal(once);
        });
    }
});
