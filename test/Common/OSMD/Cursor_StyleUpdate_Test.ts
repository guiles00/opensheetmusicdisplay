import { expect } from "chai";
import { TestUtils } from "../../Util/TestUtils";
import { OpenSheetMusicDisplay } from "../../../src/OpenSheetMusicDisplay/OpenSheetMusicDisplay";
import { Cursor } from "../../../src/OpenSheetMusicDisplay/Cursor";
import { CursorType } from "../../../src/OpenSheetMusicDisplay/OSMDOptions";

describe("Cursor style updates", () => {
    let container: HTMLElement;
    beforeEach(() => {
        container = document.createElement("div");
        container.style.width = "1300px";
        document.body.appendChild(container);
    });
    afterEach(() => {
        container.remove();
    });

    async function showCursor(type: CursorType = CursorType.Standard):
        Promise<{ osmd: OpenSheetMusicDisplay, cursor: Cursor, redrawWidths: number[] }> {
        const osmd: OpenSheetMusicDisplay = TestUtils.createOpenSheetMusicDisplay(container);
        await osmd.load(TestUtils.getScore("MuzioClementi_SonatinaOpus36No1_Part1.xml"));
        osmd.render();
        const cursor: Cursor = osmd.cursor;
        cursor.CursorOptions.type = type;
        cursor.show();
        expect(cursor.cursorElement.tagName).to.equal("DIV");

        const redrawWidths: number[] = [];
        const stockUpdateStyle: (width: number, ...rest: unknown[]) => void = (cursor as any).updateStyle;
        (cursor as any).updateStyle = function (width: number, ...rest: unknown[]): void {
            redrawWidths.push(width);
            stockUpdateStyle.call(this, width, ...rest);
        };
        return { osmd, cursor, redrawWidths };
    }

    it("doesn't reset CSS styling when stepping with unchanged options", async () => {
        const { cursor, redrawWidths } = await showCursor();
        const background: string = cursor.cursorElement.style.background;
        for (let i: number = 0; i < 50; i++) {
            cursor.next();
        }
        expect(redrawWidths.length).to.equal(0);
        expect(cursor.cursorElement.style.background).to.equal(background);
    });

    it("updates styling when an option is changed in place (#1519)", async () => {
        const { cursor, redrawWidths } = await showCursor();
        const changes: [string, () => void][] = [
            ["color", (): void => { cursor.CursorOptions.color = "#ff0000"; }],
            ["alpha", (): void => { cursor.CursorOptions.alpha = 0.9; }],
            // ThinLeft and ShortThinTopLeft have the same width, so only the type changes
            ["type (ThinLeft)", (): void => { cursor.CursorOptions.type = CursorType.ThinLeft; }],
            ["type (ShortThinTopLeft)", (): void => { cursor.CursorOptions.type = CursorType.ShortThinTopLeft; }],
        ];
        for (const [name, change] of changes) {
            const before: number = redrawWidths.length;
            change();
            cursor.update();
            expect(redrawWidths.length - before, `redraws after changing ${name}`).to.equal(1);
            cursor.next();
            expect(redrawWidths.length - before, `redraws on the next step after changing ${name}`).to.equal(1);
        }
    });

    it("updates area width without resetting CSS styling", async () => {
        const { cursor, redrawWidths } = await showCursor(CursorType.CurrentArea);
        const widths: Set<string> = new Set();
        for (let i: number = 0; i < 100; i++) {
            cursor.next();
            widths.add(cursor.cursorElement.style.width);
        }
        expect(widths.size, "the sample has measures of different widths").to.be.greaterThan(1);
        expect(redrawWidths.length, "CSS controls the appearance independently of width").to.equal(0);
        expect(cursor.cursorElement.tagName).to.equal("DIV");
    });

    it("updates the standard cursor width on zoom without resetting CSS styling", async () => {
        const { osmd, cursor, redrawWidths } = await showCursor();
        osmd.zoom = 1.5;
        cursor.update();
        expect(redrawWidths).to.deep.equal([]);
        expect(cursor.cursorElement.style.width).to.equal("45px");
    });
});
