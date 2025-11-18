import Vex from "vexflow";
import VF = Vex.Flow;
import { ColoringOptions, GraphicalNote, VisibilityOptions } from "../GraphicalNote";
import {Note} from "../../VoiceData/Note";
import {ClefInstruction} from "../../VoiceData/Instructions/ClefInstruction";
import {VexFlowConverter} from "./VexFlowConverter";
import {Pitch} from "../../../Common/DataObjects/Pitch";
import {Fraction} from "../../../Common/DataObjects/Fraction";
import {OctaveEnum, OctaveShift} from "../../VoiceData/Expressions/ContinuousExpressions/OctaveShift";
import { GraphicalVoiceEntry } from "../GraphicalVoiceEntry";
import { KeyInstruction } from "../../VoiceData/Instructions/KeyInstruction";
import { EngravingRules } from "../EngravingRules";
import { VexFlowMultiRestMeasure } from "./VexFlowMultiRestMeasure";
import { unitInPixels } from "./VexFlowMusicSheetDrawer";
import { BoundingBox } from "../BoundingBox";

/**
 * The VexFlow version of a [[GraphicalNote]].
 */
export class VexFlowGraphicalNote extends GraphicalNote {
    constructor(note: Note, parent: GraphicalVoiceEntry, activeClef: ClefInstruction,
                octaveShift: OctaveEnum = OctaveEnum.NONE, rules: EngravingRules,
                graphicalNoteLength: Fraction = undefined) {
        super(note, parent, rules, graphicalNoteLength);
        this.clef = activeClef;
        this.octaveShift = octaveShift;
        if (note.Pitch) {
            // TODO: Maybe shift to Transpose function when available
            const drawPitch: Pitch = note.isRest() ? note.Pitch : OctaveShift.getPitchFromOctaveShift(note.Pitch, octaveShift);
            this.vfpitch = VexFlowConverter.pitch(drawPitch, note.isRest(), this.clef, this.sourceNote.Notehead);
            this.vfpitch[1] = undefined;
        }
    }

    public octaveShift: OctaveEnum;
    // The pitch of this note as given by VexFlowConverter.pitch
    public vfpitch: [string, string, ClefInstruction];
    // The corresponding VexFlow StaveNote (plus its index in the chord)
    public vfnote: [VF.StemmableNote, number];
    public vfnoteIndex: number;
    // The current clef
    private clef: ClefInstruction;

    /**
     * Update the pitch of this note. Necessary in order to display accidentals correctly.
     * This is called by VexFlowGraphicalSymbolFactory.addGraphicalAccidental.
     * @param pitch
     */
    public setAccidental(pitch: Pitch): void {
        // if (this.vfnote) {
        //     let pitchAcc: AccidentalEnum = pitch.Accidental;
        //     const acc: string = Pitch.accidentalVexflow(pitch.Accidental);
        //     if (acc) {
        //         alert(acc);
        //         this.vfnote[0].addAccidental(this.vfnote[1], new VF.Accidental(acc));
        //     }
        // } else {
        // revert octave shift, as the placement of the note is independent of octave brackets
        const drawPitch: Pitch = this.drawPitch(pitch);
        // recalculate the pitch, and this time don't ignore the accidental:
        this.vfpitch = VexFlowConverter.pitch(drawPitch, this.sourceNote.isRest(), this.clef, this.sourceNote.Notehead);
        this.DrawnAccidental = drawPitch.Accidental;
        //}
    }

    public drawPitch(pitch: Pitch): Pitch {
        return OctaveShift.getPitchFromOctaveShift(pitch, this.octaveShift);
    }

    public Transpose(keyInstruction: KeyInstruction, activeClef: ClefInstruction, halfTones: number, octaveEnum: OctaveEnum): Pitch {
        const tranposedPitch: Pitch = super.Transpose(keyInstruction, activeClef, halfTones, octaveEnum);
        const drawPitch: Pitch = OctaveShift.getPitchFromOctaveShift(tranposedPitch, this.octaveShift);
        this.vfpitch = VexFlowConverter.pitch(drawPitch, this.sourceNote.isRest(), this.clef, this.sourceNote.Notehead);
        this.vfpitch[1] = undefined;
        return drawPitch;
    }

    /**
     * Set the VexFlow StaveNote corresponding to this GraphicalNote, together with its index in the chord.
     * @param note
     * @param index
     */
    public setIndex(note: VF.StemmableNote, index: number): void {
        this.vfnote = [note, index];
        this.vfnoteIndex = index;
    }

    public notehead(vfNote: VF.StemmableNote = undefined): {line: number} {
        let vfnote: any = vfNote;
        if (!vfnote) {
            vfnote = (this.vfnote[0] as any);
        }
        const noteheads: any = vfnote.note_heads;
        if (noteheads && noteheads.length > this.vfnoteIndex && noteheads[this.vfnoteIndex]) {
            return vfnote.note_heads[this.vfnoteIndex];
        } else {
            return { line: 0 };
        }
    }

    /**
     * Gets the clef for this note
     */
    public Clef(): ClefInstruction {
        return this.clef;
    }

    /**
     * Gets the id of the SVGGElement containing this note, given the SVGRenderer is used.
     * This is for low-level rendering hacks and should be used with caution.
     */
    public getSVGId(): string {
        if (!this.vfnote) {
            return undefined; // e.g. MultiRestMeasure
        }
        return this.vfnote[0].getAttribute("id");
    }

    /** Toggle visibility of the note, making it and its stem and beams invisible for `false`.
     * By default, this will also hide the note's slurs and ties (see visibilityOptions).
     * (This only works with the default SVG backend, not with the Canvas backend/renderer)
     * To get a GraphicalNote from a Note, use osmd.EngravingRules.GNote(note).
     */
    public setVisible(visible: boolean, visibilityOptions: VisibilityOptions = {}): void {
        const applyToBeams: boolean = visibilityOptions.applyToBeams ?? true; // default option if not given
        const applyToLedgerLines: boolean = visibilityOptions.applyToLedgerLines ?? true;
        const applyToNotehead: boolean = visibilityOptions.applyToNotehead ?? true;
        const applyToSlurs: boolean = visibilityOptions.applyToSlurs ?? true;
        const applyToStem: boolean = visibilityOptions.applyToStem ?? true;
        const applyToTies: boolean = visibilityOptions.applyToTies ?? true;

        const visibilityAttribute: string = "visibility";
        const visibilityString: string = visible ? "visible" : "hidden";
        if (applyToNotehead) {
            // so that it also matches undefined (option not set).
            this.getSVGGElement()?.setAttribute(visibilityAttribute, visibilityString);
        }
        // instead of setAttribute, remove() also works, but isn't reversible.
        if (applyToStem) {
            this.getStemSVG()?.setAttribute(visibilityAttribute, visibilityString);
        }
        if (applyToBeams) {
            for (const beamSVG of this.getBeamSVGs()) {
                beamSVG?.setAttribute(visibilityAttribute, visibilityString);
            }
        }
        if (applyToLedgerLines) {
            for (const ledgerSVG of this.getLedgerLineSVGs()) {
                ledgerSVG?.setAttribute(visibilityAttribute, visibilityString);
            }
        }
        if (applyToTies) {
            for (const tie of this.getTieSVGs()) {
                tie?.setAttribute(visibilityAttribute, visibilityString);
            }
        }
        if (applyToSlurs) {
            for (const slur of this.getSlurSVGs()) {
                slur?.setAttribute(visibilityAttribute, visibilityString);
            }
        }

        // usage example:
        // let voice = osmd.Sheet.Instruments[0].Voices[0];
        // for (const ve of voice.VoiceEntries) {
        //     for (const note of ve.Notes) {
        //         const gNote = osmd.EngravingRules.GNote(note);
        //         gNote?.setVisible(false);
        //     }
        // }
        // this works similarly without SVG, but with needing to render again (thus not preferable):
        // let voice = osmd.Sheet.Instruments[0].Voices[0];
        // for (const ve of voice.VoiceEntries) {
        //     for (const note of ve.Notes) {
        //         note.PrintObject = false;
        //     }
        // }
        // osmd.render();
        // this currently also still leaves ledger lines visible.
    }

    /**
     * Gets the SVGGElement containing this note, given the SVGRenderer is used.
     * This is for low-level rendering hacks and should be used with caution.
     */
    public getSVGGElement(): SVGGElement {
        if (!this.vfnote) {
            return undefined; // e.g. MultiRestMeasure
        }
        return this.vfnote[0].getAttribute("el");
    }

    private hoverHandler?: (event: MouseEvent) => void;
    private clickHandler?: (event: MouseEvent) => void;
    private mouseLeaveHandler?: (event: MouseEvent) => void;
    private interactionOverlayRect?: SVGRectElement; // Shared overlay for both hover and click

    public setHoverHandler(handler: (event: MouseEvent) => void): void {
        const svgElement: SVGGElement = this.getSVGGElement();
        if (!svgElement) {
            return;
        }
        this.removeHoverHandler();
        this.hoverHandler = handler;

        // Use shared interaction overlay if it exists (created by click handler)
        if (this.interactionOverlayRect) {
            this.interactionOverlayRect.addEventListener("mouseenter", this.hoverHandler);
            this.interactionOverlayRect.addEventListener("mousemove", this.hoverHandler);
            this.mouseLeaveHandler = (event: MouseEvent): void => {
                // Call the handler with mouseleave event so tooltip can be hidden
                // The handler should manage its own hide logic, but we pass the event
                handler(event);
            };
            this.interactionOverlayRect.addEventListener("mouseleave", this.mouseLeaveHandler);
            this.interactionOverlayRect.style.cursor = "pointer";
            return;
        }

        // Create hover overlay if no shared overlay exists
        const staffEntryBBox: BoundingBox = this.parentVoiceEntry.parentStaffEntry.PositionAndShape;
        if (staffEntryBBox) {
            staffEntryBBox.calculateAbsolutePosition();

            let svgRoot: SVGSVGElement = svgElement.ownerSVGElement;
            if (!svgRoot) {
                let parent: Node = svgElement.parentNode;
                while (parent && parent.nodeName !== "svg") {
                    parent = parent.parentNode;
                }
                svgRoot = parent as SVGSVGElement;
            }

            if (svgRoot) {
                const padding: number = 1.0;
                const leftOSMD: number = staffEntryBBox.AbsolutePosition.x + staffEntryBBox.BorderLeft - padding;
                const topOSMD: number = staffEntryBBox.AbsolutePosition.y + staffEntryBBox.BorderTop - padding;
                const widthOSMD: number = (staffEntryBBox.BorderRight - staffEntryBBox.BorderLeft) + (padding * 2);
                const heightOSMD: number = (staffEntryBBox.BorderBottom - staffEntryBBox.BorderTop) + (padding * 2);

                const left: number = leftOSMD * unitInPixels;
                const top: number = topOSMD * unitInPixels;
                const width: number = widthOSMD * unitInPixels;
                const height: number = heightOSMD * unitInPixels;

                if (width > 0 && height > 0) {
                    // Create shared interaction overlay rectangle
                    this.interactionOverlayRect = document.createElementNS("http://www.w3.org/2000/svg", "rect");
                    this.interactionOverlayRect.setAttribute("x", left.toString());
                    this.interactionOverlayRect.setAttribute("y", top.toString());
                    this.interactionOverlayRect.setAttribute("width", width.toString());
                    this.interactionOverlayRect.setAttribute("height", height.toString());
                    this.interactionOverlayRect.setAttribute("fill", "transparent");
                    this.interactionOverlayRect.setAttribute("stroke", "none");
                    this.interactionOverlayRect.setAttribute("pointer-events", "all");
                    this.interactionOverlayRect.style.cursor = "pointer";
                    this.interactionOverlayRect.setAttribute("class", "osmd-note-interaction-overlay");

                    svgRoot.appendChild(this.interactionOverlayRect);

                    // Attach hover handlers to overlay
                    this.interactionOverlayRect.addEventListener("mouseenter", this.hoverHandler);
                    this.interactionOverlayRect.addEventListener("mousemove", this.hoverHandler);
                    this.mouseLeaveHandler = (event: MouseEvent): void => {
                        // Call the handler with mouseleave event so tooltip can be hidden
                        handler(event);
                    };
                    this.interactionOverlayRect.addEventListener("mouseleave", this.mouseLeaveHandler);
                    return;
                }
            }
        }

        // Fallback to original behavior
        svgElement.style.cursor = "pointer";
        svgElement.addEventListener("mouseenter", this.hoverHandler);
        svgElement.addEventListener("mousemove", this.hoverHandler);
        this.mouseLeaveHandler = (event: MouseEvent): void => {
            // Call the handler with mouseleave event so tooltip can be hidden
            handler(event);
        };
        svgElement.addEventListener("mouseleave", this.mouseLeaveHandler);
    }

    public removeHoverHandler(): void {
        // Remove hover handlers from shared interaction overlay if it exists
        if (this.interactionOverlayRect && this.hoverHandler) {
            this.interactionOverlayRect.removeEventListener("mouseenter", this.hoverHandler);
            this.interactionOverlayRect.removeEventListener("mousemove", this.hoverHandler);
            if (this.mouseLeaveHandler) {
                this.interactionOverlayRect.removeEventListener("mouseleave", this.mouseLeaveHandler);
            }
            // Only remove cursor style if no click handler
            if (!this.clickHandler) {
                this.interactionOverlayRect.style.cursor = "";
            }
        }

        // Also remove from the original SVG element as fallback
        const svgElement: SVGGElement = this.getSVGGElement();
        if (svgElement && this.hoverHandler) {
            svgElement.removeEventListener("mouseenter", this.hoverHandler);
            svgElement.removeEventListener("mousemove", this.hoverHandler);
            if (this.mouseLeaveHandler) {
                svgElement.removeEventListener("mouseleave", this.mouseLeaveHandler);
            }
            if (!this.clickHandler) {
                svgElement.style.cursor = "";
            }
        }
        this.hoverHandler = undefined;
        this.mouseLeaveHandler = undefined;
    }

    public setClickHandler(handler: (event: MouseEvent) => void): void {
        const svgElement: SVGGElement = this.getSVGGElement();
        if (!svgElement) {
            return;
        }
        this.removeClickHandler();
        this.clickHandler = handler;

        // Create an invisible overlay rectangle using the parent StaffEntry's bounding box
        // This makes the click area larger and easier to interact with (StaffEntry is larger than VoiceEntry)
        const staffEntryBBox: BoundingBox = this.parentVoiceEntry.parentStaffEntry.PositionAndShape;
        if (staffEntryBBox) {
            // Ensure absolute positions are calculated
            staffEntryBBox.calculateAbsolutePosition();

            // Get the SVG root element (usually the parent of the note's SVG group)
            let svgRoot: SVGSVGElement = svgElement.ownerSVGElement;
            if (!svgRoot) {
                // Fallback: find the SVG root by traversing up the DOM
                let parent: Node = svgElement.parentNode;
                while (parent && parent.nodeName !== "svg") {
                    parent = parent.parentNode;
                }
                svgRoot = parent as SVGSVGElement;
            }

            if (svgRoot) {
                // Calculate the bounding box dimensions in OSMD units using StaffEntry (larger than VoiceEntry)
                // Add padding to make the click area larger and easier to interact with
                const padding: number = 1.0; // Padding in OSMD units (will be 10 pixels)
                const leftOSMD: number = staffEntryBBox.AbsolutePosition.x + staffEntryBBox.BorderLeft - padding;
                const topOSMD: number = staffEntryBBox.AbsolutePosition.y + staffEntryBBox.BorderTop - padding;
                const widthOSMD: number = (staffEntryBBox.BorderRight - staffEntryBBox.BorderLeft) + (padding * 2);
                const heightOSMD: number = (staffEntryBBox.BorderBottom - staffEntryBBox.BorderTop) + (padding * 2);

                // Convert OSMD units to SVG pixels (1 OSMD unit = unitInPixels pixels)
                const left: number = leftOSMD * unitInPixels;
                const top: number = topOSMD * unitInPixels;
                const width: number = widthOSMD * unitInPixels;
                const height: number = heightOSMD * unitInPixels;

                // Only create overlay if dimensions are valid
                if (width > 0 && height > 0) {
                    // Create or reuse shared interaction overlay rectangle
                    if (!this.interactionOverlayRect) {
                        this.interactionOverlayRect = document.createElementNS("http://www.w3.org/2000/svg", "rect");
                        this.interactionOverlayRect.setAttribute("x", left.toString());
                        this.interactionOverlayRect.setAttribute("y", top.toString());
                        this.interactionOverlayRect.setAttribute("width", width.toString());
                        this.interactionOverlayRect.setAttribute("height", height.toString());
                        this.interactionOverlayRect.setAttribute("fill", "transparent");
                        this.interactionOverlayRect.setAttribute("stroke", "none");
                        this.interactionOverlayRect.setAttribute("pointer-events", "all");
                        this.interactionOverlayRect.style.cursor = "pointer";
                        this.interactionOverlayRect.setAttribute("class", "osmd-note-interaction-overlay");

                    // Insert the overlay at the end of the SVG so it's on top (SVG uses document order for layering)
                        svgRoot.appendChild(this.interactionOverlayRect);
                    }

                    // Attach click handler to the shared overlay
                    this.interactionOverlayRect.addEventListener("click", this.clickHandler);
                } else {
                    // Fallback to original behavior if bounding box is invalid
                    svgElement.addEventListener("click", this.clickHandler);
                    svgElement.style.cursor = "pointer";
                }
            } else {
                // Fallback to original behavior if we can't find SVG root
                svgElement.addEventListener("click", this.clickHandler);
                svgElement.style.cursor = "pointer";
            }
        } else {
        // Fallback to original behavior if no StaffEntry bounding box
            svgElement.addEventListener("click", this.clickHandler);
            svgElement.style.cursor = "pointer";
        }
    }

    public removeClickHandler(): void {
        // Remove click handler from shared interaction overlay if it exists
        if (this.interactionOverlayRect && this.clickHandler) {
            this.interactionOverlayRect.removeEventListener("click", this.clickHandler);
            // Only remove overlay if no hover handler is using it
            if (!this.hoverHandler) {
                this.interactionOverlayRect.remove();
                this.interactionOverlayRect = undefined;
            } else {
                // Keep overlay but remove cursor style if no click handler
                this.interactionOverlayRect.style.cursor = "";
            }
        }

        // Also remove from the original SVG element as fallback
        const svgElement: SVGGElement = this.getSVGGElement();
        if (svgElement && this.clickHandler) {
            svgElement.removeEventListener("click", this.clickHandler);
            if (!this.hoverHandler) {
                svgElement.style.cursor = "";
            }
        }
        this.clickHandler = undefined;
    }

    /** Gets the SVG path element of the note's stem. */
    public getStemSVG(): HTMLElement {
        const groupOrPath: HTMLElement = document.getElementById("vf-" + this.getSVGId() + "-stem");
        // whether we get the group node or path node depends on whether the note has a beam, for some reason (TODO)
        if (groupOrPath?.children.length > 0) {
            return groupOrPath.children[0] as HTMLElement;
            // We want to return the same type of node every time, not a path node if no beam and a group node if it has a beam.
        }
        return groupOrPath;
        // more correct, but Vex.Prefix() is not in the definitions:
        //return document.getElementById((Vex as any).Prefix(this.getSVGId() + "-stem"));
    }

    /** Gets the SVG path elements of the beams starting on this note. */
    public getBeamSVGs(): HTMLElement[] {
        const beamSVGs: HTMLElement[] = [];
        for (let i: number = 0;; i++) {
            const newSVG: HTMLElement = document.getElementById(`vf-${this.getSVGId()}-beam${i}`);
            if (!newSVG) {
                break;
            }
            beamSVGs.push(newSVG);
        }
        return beamSVGs;
    }

    /** Gets the SVG path elements of the note's ledger lines. */
    public getLedgerLineSVGs(): HTMLElement[] {
        const ledgerSVGs: HTMLElement[] = [];
        const idString: string = `vf-${this.getSVGId()}ledgers`;
        const groupSVG: HTMLElement = document.getElementById(idString);
        if (!groupSVG) {
            return [];
        }
        for (const child of groupSVG.childNodes) {
            ledgerSVGs.push(child as HTMLElement);
        }
        return ledgerSVGs;
    }

    /** Gets the SVG path elements of the note's tie curves. */
    public getTieSVGs(): HTMLElement[] {
        const tieSVGs: HTMLElement[] = [];
        const ties: NodeListOf<HTMLElement> = document.querySelectorAll(`[id='vf-${this.getSVGId()}-tie']`);
        // TODO multiple ties have the same id sometimes, DOM elements are not supposed to have the same id, this is invalid HTML. But it works.
        for (const tie of ties) {
            tieSVGs.push(tie);
        }
        return tieSVGs;
    }

    /** Gets the SVG path elements of the note's slur curve. */
    public getSlurSVGs(): HTMLElement[] {
        const slurSVGs: HTMLElement[] = [];
        const slurs: NodeListOf<HTMLElement> = document.querySelectorAll(`[id='vf-${this.getSVGId()}-slur']`);
        // TODO multiple slurs have the same id sometimes, DOM elements are not supposed to have the same id, this is invalid HTML. But it works.
        for (const slur of slurs) {
            slurSVGs.push(slur);
        }
        return slurSVGs;
    }

    public getNoteheadSVGs(): HTMLElement[] {
        const vfNote: HTMLElement = this.getVFNoteSVG();
        const noteheads: HTMLElement[] = [];
        if (vfNote?.children?.length) {
            for (const noteChild of Array.from(vfNote.children)) {
                if (noteChild.classList.contains("vf-notehead")) {
                    noteheads.push(noteChild as HTMLElement);
                }
            }
        }
        return noteheads;
    }

    public getFlagSVG(): HTMLElement {
        const vfNote: HTMLElement = this.getVFNoteSVG();
        if (vfNote?.children?.length) {
            for (const noteChild of Array.from(vfNote.children)) {
                if (noteChild.classList.contains("vf-flag")) {
                    return noteChild as HTMLElement;
                }
            }
        }
        return undefined;
    }

    public getVFNoteSVG(): HTMLElement {
        if (this.parentVoiceEntry.parentStaffEntry.parentMeasure.isMultiRestMeasure()) {
            return this.parentVoiceEntry.parentStaffEntry.parentMeasure.multiRestElement;
        }
        const note: SVGGElement = this.getSVGGElement();
        if (note?.children) {
            for (const noteChild of Array.from(note.children)) {
                if (noteChild.classList.contains("vf-note")) {
                    return noteChild as HTMLElement;
                }
            }
        }
        return undefined;
    }

    public getModifierSVGs(): HTMLElement[] {
        const stavenote: SVGGElement = this.getSVGGElement();
        const modifierSVGs: HTMLElement[] = [];
        if (stavenote?.children?.length) {
            for (const noteChild of Array.from(stavenote.children)) {
                if (noteChild.classList.contains("vf-modifiers")) {
                    modifierSVGs.push(noteChild as HTMLElement);
                }
            }
        }
        return modifierSVGs;
    }

    /** Change the color of a note (without re-rendering). See ColoringOptions for options like applyToBeams etc.
     * This requires the SVG backend (default, instead of canvas backend).
     */
    public setColor(color: string, coloringOptions: ColoringOptions = {}): void {
        const applyToBeams: boolean = coloringOptions.applyToBeams ?? false; // default if option not given
        const applyToFlag: boolean = coloringOptions.applyToFlag ?? true;
        const applyToLedgerLines: boolean = coloringOptions.applyToLedgerLines ?? false;
        const applyToLyrics: boolean = coloringOptions.applyToLyrics ?? false;
        const applyToModifiers: boolean = coloringOptions.applyToModifiers ?? true;
        const applyToNoteheads: boolean = coloringOptions.applyToNoteheads ?? true;
        const applyToSlurs: boolean = coloringOptions.applyToSlurs ?? false;
        const applyToStem: boolean = coloringOptions.applyToStem ?? true;
        const applyToTies: boolean = coloringOptions.applyToTies ?? false;
        const applyToMultiRestMeasure: boolean = coloringOptions.applyToMultiRestMeasure ?? true;
        const applyToMultiRestMeasureNumber: boolean = coloringOptions.applyToMultiRestMeasureNumber ?? true;
        const applyToMultiRestMeasureRestBar: boolean = coloringOptions.applyToMultiRestMeasureRestBar ?? true;

        if (applyToBeams) {
            const beams: HTMLElement[] = this.getBeamSVGs();
            for (const beam of beams) {
                for (const beamPath of beam.children) {
                    beamPath.setAttribute("fill", color);
                }
            }
        }

        if (applyToFlag) {
            const flag: HTMLElement = this.getFlagSVG();
            if (flag) {
                for (const flagPath of flag.children) {
                    flagPath.setAttribute("fill", color);
                }
            }
        }

        if (applyToLedgerLines) {
            const ledgerLines: HTMLElement[] = this.getLedgerLineSVGs();
            for (const line of ledgerLines) {
                line.setAttribute("stroke", color);
            }
        }

        if (applyToLyrics) {
            const lyricsNodes: HTMLElement[] = this.getLyricsSVGs();
            for (const lyricsNode of lyricsNodes) {
                for (const textNode of lyricsNode.children) {
                    textNode.setAttribute("stroke", color);
                    textNode.setAttribute("fill", color);
                }
            }
        }

        if (applyToMultiRestMeasure) {
            if (this.parentVoiceEntry.parentStaffEntry.parentMeasure instanceof VexFlowMultiRestMeasure) {
                const measure: VexFlowMultiRestMeasure = this.parentVoiceEntry.parentStaffEntry.parentMeasure as VexFlowMultiRestMeasure;
                const svgElement: SVGGElement = measure.multiRestElementSVG;
                if (svgElement?.children?.length) {
                    for (const child of svgElement.children) { // wide bar, and number above
                        if (child.children?.length) { // vf-timesignature (multi-rest number)
                            if (applyToMultiRestMeasureNumber) {
                                for (const subChild of child.children) {
                                    subChild.setAttribute("fill", color);
                                }
                            }
                        } else {
                            if (applyToMultiRestMeasureRestBar) {
                                child.setAttribute("fill", color);
                            }
                        }
                    }
                }
            }
        }

        if (applyToModifiers) { // e.g. accidentals
            const modifiers: HTMLElement[] = this.getModifierSVGs();
            for (const modifier of modifiers) {
                for (const path of modifier.children) {
                    path.setAttribute("fill", color);
                }
            }
        }

        if (applyToNoteheads) {
            const noteheads: HTMLElement[] = this.getNoteheadSVGs();
            for (const notehead of noteheads) {
                for (const noteheadPath of notehead.children) {
                    noteheadPath.setAttribute("fill", color);
                }
            }
        }

        if (applyToSlurs) {
            const slurs: HTMLElement[] = this.getSlurSVGs();
            for (const slur of slurs) {
                for (const path of slur.children) {
                    path.setAttribute("fill", color);
                }
            }
        }

        if (applyToStem) {
            const stem: HTMLElement = this.getStemSVG();
            if (stem) {
                stem.setAttribute("stroke", color);
            }
        }

        if (applyToTies) {
            const ties: HTMLElement[] = this.getTieSVGs();
            for (const tie of ties) {
                for (const path of tie.children) {
                    path.setAttribute("fill", color);
                }
            }
        }
    }

    public setOpacity(opacity: number): void {
        const beams: HTMLElement[] = this.getBeamSVGs();
        for (const beam of beams) {
            for (const beamPath of beam.children) {
                beamPath.setAttribute("opacity", opacity.toString());
            }
        }

        const flag: HTMLElement = this.getFlagSVG();
        if (flag) {
            for (const flagPath of flag.children) {
                flagPath.setAttribute("opacity", opacity.toString());
            }
        }

        const ledgerLines: HTMLElement[] = this.getLedgerLineSVGs();
        for (const line of ledgerLines) {
            line.setAttribute("opacity", opacity.toString());
        }

        if (this.parentVoiceEntry.parentStaffEntry.parentMeasure instanceof VexFlowMultiRestMeasure) {
            const measure: VexFlowMultiRestMeasure = this.parentVoiceEntry.parentStaffEntry.parentMeasure as VexFlowMultiRestMeasure;
            const svgElement: SVGGElement = measure.multiRestElementSVG;
            if (svgElement?.children?.length) {
                for (const child of svgElement.children) {
                    if (child.children?.length) {
                        for (const subChild of child.children) {
                            subChild.setAttribute("opacity", opacity.toString());
                        }
                    } else {
                        child.setAttribute("opacity", opacity.toString());
                    }
                }
            }
        }

        const modifiers: HTMLElement[] = this.getModifierSVGs();
        for (const modifier of modifiers) {
            for (const path of modifier.children) {
                path.setAttribute("opacity", opacity.toString());
            }
        }

        const noteheads: HTMLElement[] = this.getNoteheadSVGs();
        for (const notehead of noteheads) {
            for (const noteheadPath of notehead.children) {
                noteheadPath.setAttribute("opacity", opacity.toString());
            }
        }

        const slurs: HTMLElement[] = this.getSlurSVGs();
        for (const slur of slurs) {
            for (const path of slur.children) {
                path.setAttribute("opacity", opacity.toString());
            }
        }

        const stem: HTMLElement = this.getStemSVG();
        if (stem) {
            stem.setAttribute("opacity", opacity.toString());
        }

        const ties: HTMLElement[] = this.getTieSVGs();
        for (const tie of ties) {
            for (const path of tie.children) {
                path.setAttribute("opacity", opacity.toString());
            }
        }
    }
}
