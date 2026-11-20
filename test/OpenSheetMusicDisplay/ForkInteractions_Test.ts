import { expect } from "vitest";
import { OpenSheetMusicDisplay } from "../../src/OpenSheetMusicDisplay/OpenSheetMusicDisplay";
import { VexFlowGraphicalNote } from "../../src/MusicalScore/Graphical/VexFlow/VexFlowGraphicalNote";
import { GraphicalChordSymbolContainer } from "../../src/MusicalScore/Graphical/GraphicalChordSymbolContainer";
import { GraphicalLabel } from "../../src/MusicalScore/Graphical/GraphicalLabel";
import { Fraction } from "../../src/Common/DataObjects/Fraction";
import { TestUtils } from "../Util/TestUtils";

/**
 * Fork (Piano Tree) APIs the app relies on for interaction, which upstream OSMD doesn't have: hover and click handlers
 * on notes and labels, the handlers for a note's SVG being released and drawn again, the list of all chord symbols,
 * and the octave shift that is active at a position.
 */
describe("Fork interaction APIs", () => {
    let div: HTMLElement;
    let osmd: OpenSheetMusicDisplay;

    afterEach(() => {
        osmd?.clear();
        div?.remove();
    });

    async function render(score: string | Document): Promise<OpenSheetMusicDisplay> {
        div = TestUtils.getDivElement(document);
        div.style.width = "1000px";
        osmd = new OpenSheetMusicDisplay(div, { autoResize: false, backend: "svg" });
        await osmd.load(score);
        osmd.render();
        return osmd;
    }

    function firstNote(): VexFlowGraphicalNote {
        return osmd.GraphicSheet.MeasureList[0][0].staffEntries[0].graphicalVoiceEntries[0].notes[0] as VexFlowGraphicalNote;
    }

    function fire(element: Element, type: string): void {
        element.dispatchEvent(new MouseEvent(type, { bubbles: true }));
    }

    it("calls a note's click handler until it is removed, and shows a pointer meanwhile", async () => {
        await render(TestUtils.getScore("MuzioClementi_SonatinaOpus36No1_Part1.xml"));
        const note: VexFlowGraphicalNote = firstNote();
        const element: SVGGElement = note.getSVGGElement();
        let clicks: number = 0;
        note.setClickHandler(() => clicks++);
        expect(element.style.cursor).to.equal("pointer");
        fire(element, "click");
        expect(clicks).to.equal(1);
        note.removeClickHandler();
        fire(element, "click");
        expect(clicks, "clicks after the handler is removed").to.equal(1);
        expect(element.style.cursor).to.equal("");
    });

    it("calls a note's hover handler on enter and move until it is removed", async () => {
        await render(TestUtils.getScore("MuzioClementi_SonatinaOpus36No1_Part1.xml"));
        const note: VexFlowGraphicalNote = firstNote();
        const element: SVGGElement = note.getSVGGElement();
        let hovers: number = 0;
        note.setHoverHandler(() => hovers++);
        expect(element.style.cursor).to.equal("pointer");
        fire(element, "mouseenter");
        fire(element, "mousemove");
        expect(hovers).to.equal(2);
        note.removeHoverHandler();
        fire(element, "mouseenter");
        fire(element, "mousemove");
        expect(hovers, "hovers after the handler is removed").to.equal(2);
        expect(element.style.cursor).to.equal("");
    });

    it("tells when a note's SVG is released and drawn again, and puts the click handler on the new element", async () => {
        await render(TestUtils.getScore("MuzioClementi_SonatinaOpus36No1_Part1.xml"));
        const note: VexFlowGraphicalNote = firstNote();
        let materialized: number = 0;
        note.addSVGMaterializationHandler(() => materialized++);
        expect(materialized, "called right away for a note that is drawn").to.equal(1);

        let evicted: number = 0;
        let recreated: number = 0;
        let clicks: number = 0;
        note.addSVGEvictionHandler(() => evicted++);
        note.addSVGRecreationHandler(() => recreated++);
        note.setClickHandler(() => clicks++);
        const firstElement: SVGGElement = note.getSVGGElement();

        note.releaseRenderedSVG();
        expect(evicted).to.equal(1);
        expect(note.getSVGGElement(), "the element after the release").to.equal(undefined);
        note.addSVGMaterializationHandler(() => materialized++);
        expect(materialized, "waits for the note to be drawn").to.equal(1);

        osmd.render();
        const sameNote: VexFlowGraphicalNote = firstNote();
        expect(sameNote === note, "render() keeps the graphical note").to.equal(true);
        const secondElement: SVGGElement = note.getSVGGElement();
        expect(secondElement, "the element after rendering again").to.not.equal(undefined);
        expect(secondElement === firstElement, "a new element").to.equal(false);
        expect(recreated).to.be.greaterThan(0);
        expect(materialized, "called once when the note is drawn").to.equal(2);
        fire(secondElement, "click");
        expect(clicks, "the click handler on the new element").to.equal(1);

        const recreatedBefore: number = recreated;
        osmd.render();
        expect(recreated, "the recreation handler stays for later renders").to.be.greaterThan(recreatedBefore);
        expect(materialized, "the materialization handler is used once").to.equal(2);
    });

    it("lists every chord symbol, and its label takes click and hover handlers", async () => {
        const score: Document = TestUtils.getScore("test_chord_symbols_overlap_narrow_measure_1688.musicxml");
        await render(score);
        const containers: GraphicalChordSymbolContainer[] = osmd.getAllChordSymbolContainers();
        expect(containers.length).to.equal(score.querySelectorAll("harmony").length);
        expect(osmd.GraphicSheet.getAllChordSymbolContainers().length).to.equal(containers.length);

        const label: GraphicalLabel = containers[0].GraphicalLabel;
        const element: Element = label.SVGNode as Element;
        expect(element instanceof Element, "the label is drawn").to.equal(true);
        let materialized: number = 0;
        label.addSVGMaterializationHandler(() => materialized++);
        expect(materialized).to.equal(1);
        let clicks: number = 0;
        let hovers: number = 0;
        label.setClickHandler(() => clicks++);
        label.setHoverHandler(() => hovers++);
        fire(element, "click");
        fire(element, "mouseenter");
        expect([clicks, hovers]).to.deep.equal([1, 1]);
        label.removeClickHandler();
        label.removeHoverHandler();
        fire(element, "click");
        fire(element, "mouseenter");
        expect([clicks, hovers], "after the handlers are removed").to.deep.equal([1, 1]);
    });

    it("lists no chord symbols before a sheet is loaded", () => {
        div = TestUtils.getDivElement(document);
        osmd = new OpenSheetMusicDisplay(div, { autoResize: false });
        expect(osmd.getAllChordSymbolContainers()).to.deep.equal([]);
    });

    it("finds the octave shift that is active at a position", async () => {
        const note: (step: string) => string = (step: string): string =>
            `<note><pitch><step>${step}</step><octave>5</octave></pitch><duration>1</duration><voice>1</voice><type>quarter</type></note>`;
        const shift: (type: string) => string = (type: string): string =>
            `<direction placement="above"><direction-type><octave-shift type="${type}" size="8"/></direction-type></direction>`;
        const xml: string = `<?xml version="1.0" encoding="UTF-8"?>
<score-partwise version="4.0"><part-list><score-part id="P1"><part-name>Flute</part-name></score-part></part-list>
<part id="P1"><measure number="1">
<attributes><divisions>1</divisions><key><fifths>0</fifths></key><time><beats>4</beats><beat-type>4</beat-type></time>
<clef><sign>G</sign><line>2</line></clef></attributes>
${shift("down")}${note("C")}${note("D")}${note("E")}${note("F")}</measure>
<measure number="2">${note("G")}${note("A")}${shift("stop")}${note("B")}${note("C")}</measure>
<measure number="3">${note("C")}${note("D")}${note("E")}${note("F")}</measure></part></score-partwise>`;
        await render(xml);
        expect(osmd.getActiveOctaveShift(new Fraction(0, 1), 0, 0), "at the start of the shift").to.not.equal(undefined);
        expect(osmd.getActiveOctaveShift(new Fraction(3, 4), 0, 0), "inside the first measure").to.not.equal(undefined);
        expect(osmd.getActiveOctaveShift(new Fraction(0, 1), 0, 1), "at the start of the second measure").to.not.equal(undefined);
        expect(osmd.getActiveOctaveShift(new Fraction(3, 4), 0, 1), "after the stop").to.equal(undefined);
        expect(osmd.getActiveOctaveShift(new Fraction(0, 1), 0, 2), "in the measure after the shift").to.equal(undefined);
        expect(osmd.getActiveOctaveShift(new Fraction(0, 1), 1, 0), "on a staff the sheet doesn't have").to.equal(undefined);
        expect(osmd.getActiveOctaveShift(new Fraction(0, 1), 0, 9), "in a measure the sheet doesn't have").to.equal(undefined);
    });
});
