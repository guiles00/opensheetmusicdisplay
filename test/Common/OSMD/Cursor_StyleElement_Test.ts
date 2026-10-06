import { expect } from "chai";
import { TestUtils } from "../../Util/TestUtils";
import { OpenSheetMusicDisplay } from "../../../src/OpenSheetMusicDisplay/OpenSheetMusicDisplay";
import { Cursor } from "../../../src/OpenSheetMusicDisplay/Cursor";
import { CursorType } from "../../../src/OpenSheetMusicDisplay/OSMDOptions";

describe("Cursor CSS element", () => {
    let container: HTMLElement;
    let pageStyle: HTMLStyleElement;
    beforeEach(() => {
        container = document.createElement("div");
        container.style.width = "1300px";
        document.body.appendChild(container);
    });
    afterEach(() => {
        container.remove();
        pageStyle?.remove();
        pageStyle = undefined;
    });

    async function showCursor(type: CursorType = CursorType.Standard): Promise<Cursor> {
        const osmd: OpenSheetMusicDisplay = TestUtils.createOpenSheetMusicDisplay(container);
        await osmd.load(TestUtils.getScore("MuzioClementi_SonatinaOpus36No1_Part1.xml"));
        osmd.render();
        const cursor: Cursor = osmd.cursor;
        cursor.CursorOptions.type = type;
        cursor.show();
        return cursor;
    }

    it("keeps its height with page CSS that overrides the height attribute, like Tailwind's img { height: auto }", async () => {
        pageStyle = document.createElement("style");
        pageStyle.textContent = "img { max-width: 100%; height: auto; }"; // from Tailwind's base styles (preflight)
        document.head.appendChild(pageStyle);
        const cursor: Cursor = await showCursor();
        const height: number = cursor.cursorElement.getBoundingClientRect().height;
        // the height of the system, not the 1 pixel of the image
        expect(height).to.be.greaterThan(50);
        expect(height).to.be.closeTo(Number.parseFloat(cursor.cursorElement.style.height), 1);
    });

    it("keeps a height that the app set with !important", async () => {
        const cursor: Cursor = await showCursor();
        cursor.cursorElement.style.setProperty("height", "50px", "important");
        cursor.next();
        cursor.next();
        expect(cursor.cursorElement.getBoundingClientRect().height).to.equal(50);
    });

    it("keeps a div cursor that consumers can style with CSS", async () => {
        pageStyle = document.createElement("style");
        pageStyle.textContent = "[id^='cursorImg-'] { background: rgb(255, 0, 0); opacity: 0.7; }";
        document.head.appendChild(pageStyle);
        const cursor: Cursor = await showCursor();
        expect(cursor.cursorElement.tagName).to.equal("DIV");
        expect(getComputedStyle(cursor.cursorElement).backgroundColor).to.equal("rgb(255, 0, 0)");
        expect(getComputedStyle(cursor.cursorElement).opacity).to.equal("0.7");
        cursor.next();
        expect(cursor.cursorElement.style.background).to.equal("");
        expect(cursor.cursorElement.style.opacity).to.equal("");
    });

    it("is hidden from screen readers and can't be dragged", async () => {
        const cursor: Cursor = await showCursor();
        expect(cursor.cursorElement.getAttribute("aria-hidden")).to.equal("true");
        expect(cursor.cursorElement.draggable).to.equal(false);
    });
});
