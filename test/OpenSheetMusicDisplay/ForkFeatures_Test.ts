import { expect } from "chai";
import { OpenSheetMusicDisplay } from "../../src/OpenSheetMusicDisplay/OpenSheetMusicDisplay";
import { VexFlowGraphicalNote } from "../../src/MusicalScore/Graphical/VexFlow/VexFlowGraphicalNote";
import { GraphicalMusicSheet } from "../../src/MusicalScore/Graphical/GraphicalMusicSheet";
import { TestUtils } from "../Util/TestUtils";

/**
 * Fork (Piano Tree) APIs the app relies on that upstream OSMD doesn't have: voice and staff dimming for hand modes,
 * staffline lookup by staff index, MIDI export, and the deferred graphic preparation kept through upstream merges.
 */
describe("Fork features", () => {
    // A piano score: voice 1 (stems up: two beamed eighths, a flagged eighth, an eighth rest, a half note) and voice 2
    //   (stems down: a half note, a quarter, a quarter) on the treble staff, voice 5 (a whole note) on the bass staff.
    const pitch: (step: string, octave: number) => string = (step: string, octave: number): string =>
        `<pitch><step>${step}</step><octave>${octave}</octave></pitch>`;
    const pianoScore: string = `<?xml version="1.0" encoding="UTF-8"?>
<score-partwise version="4.0"><part-list><score-part id="P1"><part-name>Piano</part-name></score-part></part-list>
<part id="P1"><measure number="1">
<attributes><divisions>2</divisions><key><fifths>0</fifths></key><time><beats>4</beats><beat-type>4</beat-type></time>
<staves>2</staves><clef number="1"><sign>G</sign><line>2</line></clef><clef number="2"><sign>F</sign><line>4</line></clef></attributes>
<note>${pitch("C", 5)}<duration>1</duration><voice>1</voice><type>eighth</type><stem>up</stem><staff>1</staff><beam number="1">begin</beam></note>
<note>${pitch("D", 5)}<duration>1</duration><voice>1</voice><type>eighth</type><stem>up</stem><staff>1</staff><beam number="1">end</beam></note>
<note>${pitch("E", 5)}<duration>1</duration><voice>1</voice><type>eighth</type><stem>up</stem><staff>1</staff></note>
<note><rest/><duration>1</duration><voice>1</voice><type>eighth</type><staff>1</staff></note>
<note>${pitch("F", 5)}<duration>4</duration><voice>1</voice><type>half</type><stem>up</stem><staff>1</staff></note>
<backup><duration>8</duration></backup>
<note>${pitch("E", 4)}<duration>4</duration><voice>2</voice><type>half</type><stem>down</stem><staff>1</staff></note>
<note>${pitch("F", 4)}<duration>2</duration><voice>2</voice><type>quarter</type><stem>down</stem><staff>1</staff></note>
<note>${pitch("G", 4)}<duration>2</duration><voice>2</voice><type>quarter</type><stem>down</stem><staff>1</staff></note>
<backup><duration>8</duration></backup>
<note>${pitch("C", 3)}<duration>8</duration><voice>5</voice><type>whole</type><staff>2</staff></note>
</measure></part></score-partwise>`;

    let div: HTMLElement;
    let osmd: OpenSheetMusicDisplay;

    beforeEach(() => {
        div = TestUtils.getDivElement(document);
        div.style.width = "800px";
        osmd = new OpenSheetMusicDisplay(div, { autoResize: false, backend: "svg" });
    });

    afterEach(() => {
        osmd.dispose();
        div.remove();
    });

    function notesOfVoice(voiceId: number): VexFlowGraphicalNote[] {
        const notes: VexFlowGraphicalNote[] = [];
        for (const staff of osmd.Sheet.Staves) {
            for (const voice of staff.Voices) {
                if (voice.VoiceId !== voiceId) {
                    continue;
                }
                for (const voiceEntry of voice.VoiceEntries) {
                    for (const note of voiceEntry.Notes) {
                        if (note.Pitch) {
                            notes.push(osmd.EngravingRules.GNote(note) as VexFlowGraphicalNote);
                        }
                    }
                }
            }
        }
        return notes;
    }

    /** The opacity a shape is drawn with: set on the shape, or on each shape of its group. */
    function opacitiesOf(element: Element): string[] {
        const shapes: Element[] = element.children.length === 0 ? [element] : Array.from(element.children);
        return shapes.map(shape => shape.getAttribute("opacity") ?? "1");
    }

    /** The opacities of a note's heads, stem, beams and flag. */
    function noteOpacities(note: VexFlowGraphicalNote): string[] {
        const parts: Element[] = [...note.getNoteheadSVGs(), note.getStemSVG(), ...note.getBeamSVGs(), note.getFlagSVG()]
            .filter(part => part);
        return parts.flatMap(opacitiesOf);
    }

    for (const mergeSingleChildGroups of [true, false]) {
        describe(`with SVGMergeSingleChildGroups ${mergeSingleChildGroups}`, () => {
            beforeEach(async () => {
                osmd.EngravingRules.SVGMergeSingleChildGroups = mergeSingleChildGroups;
                await osmd.load(pianoScore);
                osmd.render();
            });

            it("blurVoices() dims the heads, stems, beams and flags of the given voices only, resetOpacity() restores them", () => {
                const upperVoice: VexFlowGraphicalNote[] = notesOfVoice(1);
                expect(upperVoice.some(note => note.getBeamSVGs().length > 0), "a beamed note in voice 1").to.equal(true);
                expect(upperVoice.some(note => note.getFlagSVG()), "a flagged note in voice 1").to.equal(true);

                osmd.blurVoices([1], 0.2);
                for (const note of upperVoice) {
                    const opacities: string[] = noteOpacities(note);
                    expect(opacities.length, "the note's shapes").to.be.greaterThan(0);
                    expect(opacities.every(opacity => opacity === "0.2"), `voice 1 dimmed: ${opacities}`).to.equal(true);
                }
                for (const note of [...notesOfVoice(2), ...notesOfVoice(5)]) {
                    expect(noteOpacities(note).every(opacity => opacity === "1"), "the other voices stay").to.equal(true);
                }

                osmd.resetOpacity();
                for (const note of [...upperVoice, ...notesOfVoice(2), ...notesOfVoice(5)]) {
                    expect(noteOpacities(note).every(opacity => opacity === "1"), "restored").to.equal(true);
                }
            });

            it("blurAllVoicesExceptVoices() dims every other voice", () => {
                osmd.blurAllVoicesExceptVoices([1, 2], 0.3);
                for (const note of [...notesOfVoice(1), ...notesOfVoice(2)]) {
                    expect(noteOpacities(note).every(opacity => opacity === "1"), "kept voices").to.equal(true);
                }
                for (const note of notesOfVoice(5)) {
                    expect(noteOpacities(note).every(opacity => opacity === "0.3"), "the left hand dimmed").to.equal(true);
                }
            });
        });
    }

    it("draws stafflines with the staffline and osmd-staffline classes and their staff index, which blurStaff() dims", async () => {
        await osmd.load(pianoScore);
        osmd.render();
        const stafflines: (staffIndex: number) => SVGGElement[] = (staffIndex: number): SVGGElement[] =>
            Array.from(div.querySelectorAll<SVGGElement>(`.staffline[data-staff-index="${staffIndex}"]`));
        expect(stafflines(0).length, "the treble staff's staffline").to.equal(1);
        expect(stafflines(1).length, "the bass staff's staffline").to.equal(1);
        expect(stafflines(1)[0].classList.contains("osmd-staffline"), "upstream's class too").to.equal(true);

        osmd.blurStaff(1, 0.3);
        expect(stafflines(1)[0].style.opacity).to.equal("0.3");
        expect(stafflines(0)[0].style.opacity, "the other staff stays").to.equal("");
        osmd.restoreStaff(1);
        expect(stafflines(1)[0].style.opacity).to.equal("");
    });

    it("keeps the graphic across renders, and creates it again only after a change of the transposition", async () => {
        await osmd.load(pianoScore);
        // the graphic is prepared on the first render, which must not take the unprepared graphic for a changed transposition
        expect(osmd.GraphicSheet.TranspositionChanged, "before the first render").to.equal(false);
        osmd.render();
        const graphic: GraphicalMusicSheet = osmd.GraphicSheet;
        osmd.render();
        expect(osmd.GraphicSheet, "a second render").to.equal(graphic);

        osmd.Sheet.Transpose = 2;
        osmd.render();
        expect(osmd.GraphicSheet, "a render after a transposition").to.not.equal(graphic);
        expect(notesOfVoice(5).every(note => note?.getSVGGElement()), "the notes drawn on the new graphic").to.equal(true);
        expect(osmd.GraphicSheet.TranspositionChanged, "after it").to.equal(false);
        const transposed: GraphicalMusicSheet = osmd.GraphicSheet;
        osmd.render();
        expect(osmd.GraphicSheet, "a render after that").to.equal(transposed);
    });

    it("exports the loaded score as a standard MIDI file", async () => {
        await osmd.load(pianoScore);
        const midi: Uint8Array = osmd.exportMIDI();
        expect(midi, "a file").to.not.equal(undefined);
        expect(String.fromCharCode(...Array.from(midi.slice(0, 4))), "the header chunk").to.equal("MThd");
        expect(String.fromCharCode(...Array.from(midi.slice(14, 18))), "a track chunk").to.equal("MTrk");
        const hasNoteOn: (midiNote: number) => boolean = (midiNote: number): boolean =>
            midi.some((byte: number, i: number): boolean => Math.floor(byte / 16) === 9 && midi[i + 1] === midiNote && midi[i + 2] > 0);
        expect(hasNoteOn(48), "the left hand's C3").to.equal(true);
        expect(hasNoteOn(72), "the right hand's C5").to.equal(true);
    });
});
