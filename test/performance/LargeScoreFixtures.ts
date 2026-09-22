export interface ILargeScoreFixtureOptions {
    measures: number;
    graceEvery?: number;
    ornamentEvery?: number;
    tieEvery?: number;
    fingeringEvery?: number;
    meterChangeEvery?: number;
    systemBreakEvery?: number;
}

interface IMeter {
    beats: number;
    beatType: number;
    eighths: number;
    lowerNotes: { duration: number, type: string, dot: boolean }[];
}

const DIVISIONS: number = 4;
const METERS: IMeter[] = [
    { beats: 4, beatType: 4, eighths: 8, lowerNotes: Array(4).fill({ duration: 4, type: "quarter", dot: false }) },
    { beats: 3, beatType: 4, eighths: 6, lowerNotes: Array(3).fill({ duration: 4, type: "quarter", dot: false }) },
    { beats: 6, beatType: 8, eighths: 6, lowerNotes: Array(2).fill({ duration: 6, type: "quarter", dot: true }) }
];
const STEPS: string[] = ["C", "D", "E", "F", "G", "A", "B"];

function pitchXml(diatonicIndex: number): string {
    const step: string = STEPS[((diatonicIndex % 7) + 7) % 7];
    const octave: number = Math.floor(diatonicIndex / 7);
    return `<pitch><step>${step}</step><octave>${octave}</octave></pitch>`;
}

function noteXml(parts: {
    diatonic: number; duration: number; type: string; voice: number; staff: number;
    chord?: boolean; dot?: boolean; beam?: string; grace?: boolean; tie?: "start" | "stop";
    ornament?: string; fingering?: number;
}): string {
    const notations: string[] = [];
    if (parts.tie) {
        notations.push(`<tied type="${parts.tie}"/>`);
    }
    if (parts.ornament) {
        notations.push(`<ornaments><${parts.ornament}/></ornaments>`);
    }
    if (parts.fingering !== undefined) {
        notations.push(`<technical><fingering>${parts.fingering}</fingering></technical>`);
    }
    return [
        "<note>",
        parts.grace ? "<grace/>" : "",
        parts.chord ? "<chord/>" : "",
        pitchXml(parts.diatonic),
        parts.grace ? "" : `<duration>${parts.duration}</duration>`,
        parts.tie ? `<tie type="${parts.tie}"/>` : "",
        `<voice>${parts.voice}</voice>`,
        `<type>${parts.type}</type>`,
        parts.dot ? "<dot/>" : "",
        `<staff>${parts.staff}</staff>`,
        parts.beam ? `<beam number="1">${parts.beam}</beam>` : "",
        notations.length ? `<notations>${notations.join("")}</notations>` : "",
        "</note>"
    ].join("");
}

/** Deterministic dense two-staff piano score for scaling benchmarks and correctness fixtures. */
export function generateLargePianoScore(options: ILargeScoreFixtureOptions): string {
    const measures: string[] = [];
    let previousMeter: IMeter | undefined;
    for (let measureIndex: number = 0; measureIndex < options.measures; measureIndex++) {
        const meterIndex: number = options.meterChangeEvery
            ? Math.floor(measureIndex / options.meterChangeEvery) % METERS.length
            : 0;
        const meter: IMeter = METERS[meterIndex];
        const content: string[] = [];
        if (options.systemBreakEvery && measureIndex > 0 && measureIndex % options.systemBreakEvery === 0) {
            content.push("<print new-system=\"yes\"/>");
        }
        if (measureIndex === 0 || meter !== previousMeter) {
            content.push("<attributes>");
            if (measureIndex === 0) {
                content.push(`<divisions>${DIVISIONS}</divisions><key><fifths>0</fifths></key>`);
            }
            content.push(`<time><beats>${meter.beats}</beats><beat-type>${meter.beatType}</beat-type></time>`);
            if (measureIndex === 0) {
                content.push("<staves>2</staves><clef number=\"1\"><sign>G</sign><line>2</line></clef>");
                content.push("<clef number=\"2\"><sign>F</sign><line>4</line></clef>");
            }
            content.push("</attributes>");
        }
        previousMeter = meter;

        const tiesIntoThisMeasure: boolean = !!options.tieEvery && measureIndex > 0 && (measureIndex - 1) % options.tieEvery === 0;
        const tiesOutOfThisMeasure: boolean = !!options.tieEvery && measureIndex % options.tieEvery === 0 &&
            measureIndex < options.measures - 1;
        for (let eighth: number = 0; eighth < meter.eighths; eighth++) {
            const diatonic: number = 30 + ((measureIndex * 3 + eighth) % 9);
            const isFirst: boolean = eighth === 0;
            const isLast: boolean = eighth === meter.eighths - 1;
            if (options.graceEvery && isFirst && measureIndex % options.graceEvery === 0) {
                content.push(noteXml({ diatonic: diatonic + 2, duration: 0, type: "16th", voice: 1, staff: 1, grace: true, beam: "begin" }));
                content.push(noteXml({ diatonic: diatonic + 1, duration: 0, type: "16th", voice: 1, staff: 1, grace: true, beam: "end" }));
            }
            const pitch: number = (isFirst && tiesIntoThisMeasure) || (isLast && tiesOutOfThisMeasure) ? 35 : diatonic;
            content.push(noteXml({
                diatonic: pitch,
                duration: 2,
                type: "eighth",
                voice: 1,
                staff: 1,
                beam: eighth % 2 === 0 ? "begin" : "end",
                tie: isFirst && tiesIntoThisMeasure ? "stop" : isLast && tiesOutOfThisMeasure ? "start" : undefined,
                ornament: options.ornamentEvery && eighth === 2 && measureIndex % options.ornamentEvery === 0
                    ? (measureIndex / options.ornamentEvery) % 2 === 0 ? "trill-mark" : "mordent"
                    : undefined,
                fingering: options.fingeringEvery && eighth % options.fingeringEvery === 0 ? (eighth % 5) + 1 : undefined
            }));
            if (eighth % 2 === 1) {
                content.push(noteXml({ diatonic: pitch + 2, duration: 2, type: "eighth", voice: 1, staff: 1, chord: true }));
            }
        }
        content.push(`<backup><duration>${meter.eighths * 2}</duration></backup>`);
        meter.lowerNotes.forEach((lower, index): void => {
            const root: number = 16 + ((measureIndex + index * 2) % 7);
            content.push(noteXml({ diatonic: root, duration: lower.duration, type: lower.type, dot: lower.dot, voice: 5, staff: 2 }));
            content.push(noteXml({
                diatonic: root + 4, duration: lower.duration, type: lower.type, dot: lower.dot, voice: 5, staff: 2, chord: true
            }));
        });
        measures.push(`<measure number="${measureIndex + 1}">${content.join("")}</measure>`);
    }
    return [
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>",
        "<score-partwise version=\"3.1\">",
        "<part-list><score-part id=\"P1\"><part-name>Piano</part-name></score-part></part-list>",
        `<part id="P1">${measures.join("")}</part>`,
        "</score-partwise>"
    ].join("");
}
