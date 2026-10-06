import {Pitch} from "../../Common/DataObjects/Pitch";
import {KeyInstruction} from "../VoiceData/Instructions/KeyInstruction";
import {EngravingRules} from "../Graphical/EngravingRules";

export interface ITransposeCalculator {
    /** Set by OpenSheetMusicDisplay.TransposeCalculator, so that a calculator can read transposition rules
     * like EngravingRules.StrictTransposeSpelling. Optional: a custom calculator can ignore the rules. */
    rules?: EngravingRules;
    /** Transposes a note's pitch, or the root or bass pitch of a chord symbol (chordSymbol: true),
     * which a calculator can spell differently from notes, e.g. without double sharps/flats. */
    transposePitch(pitch: Pitch, currentKeyInstruction: KeyInstruction, halftones: number, chordSymbol?: boolean): Pitch;
    transposeKey(keyInstruction: KeyInstruction, transpose: number): void;
}
