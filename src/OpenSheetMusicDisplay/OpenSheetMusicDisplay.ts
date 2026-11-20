import { IXmlElement, removeWhitespaceBetweenElements } from "./../Common/FileIO/Xml";
import { VexFlowMusicSheetCalculator } from "./../MusicalScore/Graphical/VexFlow/VexFlowMusicSheetCalculator";
import { VexFlowBackend } from "./../MusicalScore/Graphical/VexFlow/VexFlowBackend";
import { VexFlowTextMeasurer } from "./../MusicalScore/Graphical/VexFlow/VexFlowTextMeasurer";
import { MusicSheetReader } from "./../MusicalScore/ScoreIO/MusicSheetReader";
import { GraphicalMusicSheet } from "./../MusicalScore/Graphical/GraphicalMusicSheet";
import { MusicSheetCalculator } from "./../MusicalScore/Graphical/MusicSheetCalculator";
import { VexFlowMusicSheetDrawer } from "./../MusicalScore/Graphical/VexFlow/VexFlowMusicSheetDrawer";
import { SvgVexFlowBackend } from "./../MusicalScore/Graphical/VexFlow/SvgVexFlowBackend";
import { CanvasVexFlowBackend } from "./../MusicalScore/Graphical/VexFlow/CanvasVexFlowBackend";
import { MusicSheet } from "./../MusicalScore/MusicSheet";
import { Cursor } from "./Cursor";
import { MXLFile, MXLHelper } from "../Common/FileIO/Mxl";
import { AJAX } from "./AJAX";
import { log } from "../Common/Logging";
import { DrawingParameters } from "../MusicalScore/Graphical/DrawingParameters";
import { DrawingParametersEnum } from "../Common/Enums/DrawingParametersEnum";
import { ColoringModes } from "../Common/Enums/ColoringModes";
import { IOSMDOptions, OSMDOptions, AutoBeamOptions, BackendType, CursorOptions, CursorType, IRenderInfo } from "./OSMDOptions";
import { EngravingRules, PageFormat } from "../MusicalScore/Graphical/EngravingRules";
import { AbstractExpression, PlacementEnum } from "../MusicalScore/VoiceData/Expressions/AbstractExpression";
import { Dictionary } from "typescript-collections";
import { AutoColorSet, GraphicalLayers } from "../MusicalScore/Graphical/DrawingEnums";
import { GraphicalMusicPage } from "../MusicalScore/Graphical/GraphicalMusicPage";
import { MusicSystem } from "../MusicalScore/Graphical/MusicSystem";
import { GraphicalMeasure } from "../MusicalScore/Graphical/GraphicalMeasure";
import { SourceMeasure } from "../MusicalScore/VoiceData/SourceMeasure";
import { MusicPartManagerIterator } from "../MusicalScore/MusicParts/MusicPartManagerIterator";
import { ITransposeCalculator } from "../MusicalScore/Interfaces/ITransposeCalculator";
import { NoteEnum } from "../Common/DataObjects/Pitch";
import { GraphicalNote } from "../MusicalScore/Graphical/GraphicalNote";
import { Fraction } from "../Common/DataObjects/Fraction";
import { MultiExpression } from "../MusicalScore/VoiceData/Expressions/MultiExpression";
import { OctaveShift } from "../MusicalScore/VoiceData/Expressions/ContinuousExpressions/OctaveShift";
import { GraphicalLyricEntry } from "../MusicalScore/Graphical/GraphicalLyricEntry";
import { GraphicalChordSymbolContainer } from "../MusicalScore/Graphical/GraphicalChordSymbolContainer";
import { GraphicalLabel } from "../MusicalScore/Graphical/GraphicalLabel";
import { MultiTempoExpression } from "../MusicalScore/VoiceData/Expressions/MultiTempoExpression";
import { MidiExporter, MidiExportOptions } from "../MusicalScore/Export/MidiExporter";
import { PointF2D } from "../Common/DataObjects/PointF2D";
import { RectangleF2D } from "../Common/DataObjects/RectangleF2D";
import { tempoLabelFromBpm } from "../Common/Tempo/TempoLabelFromBpm";
import { GraphicalStaffEntry } from "../MusicalScore/Graphical/GraphicalStaffEntry";
import { VexFlowGraphicalNote } from "../MusicalScore/Graphical/VexFlow/VexFlowGraphicalNote";
import { VerticalGraphicalStaffEntryContainer } from "../MusicalScore/Graphical/VerticalGraphicalStaffEntryContainer";
import { AbstractGraphicalExpression } from "../MusicalScore/Graphical/AbstractGraphicalExpression";
import { countLedgerLineNotesForTransposition as countLedgerLineNotesOnMusicSheet } from "./ledgerLineTranspositionCount";
import { RangeSelectionElementCollector } from "./RangeSelectionElementCollector";
import {
    createResolvedRangeSelectionConfig,
    ResolvedRangeSelectionConfig,
    RangeSelectionAnchor,
    RangeSelectionDirection,
    RangeSelectionPayload
} from "./RangeSelection";
import { TemposCalculator } from "../MusicalScore/ScoreIO/MusicSymbolModules/TemposCalculator";
import { SourceStaffEntry } from "../MusicalScore/VoiceData/SourceStaffEntry";
import { TechnicalInstruction, TechnicalInstructionType } from "../MusicalScore/VoiceData/Instructions/TechnicalInstruction";
import { Label } from "../MusicalScore/Label";
import { TextAlignmentEnum } from "../Common/Enums/TextAlignment";
import { VoiceEntry } from "../MusicalScore/VoiceData/VoiceEntry";
import { Note } from "../MusicalScore/VoiceData/Note";
import { BoundingBox } from "../MusicalScore/Graphical/BoundingBox";
import { StaffLine } from "../MusicalScore/Graphical/StaffLine";
import { SkyBottomLineCalculator } from "../MusicalScore/Graphical/SkyBottomLineCalculator";
import {
    ISystemVirtualizationOptions,
    ISystemVirtualizationStats,
    ISystemLifecycleEvent,
    SystemLifecycleListener,
    IVirtualSystemDescriptor,
    SystemVirtualizationController
} from "./SystemVirtualizationController";
import { WarmKeeperScoreMxlBase64 } from "./WarmKeeperScore";
import { CollectionUtil } from "../Util/CollectionUtil";

/** WeakRef (ES2021) isn't in the TypeScript lib of OSMD's target, es2020. Older browsers don't have it: Chrome before 84,
 *  Firefox before 79, Safari before 14.1. */
declare const WeakRef: new <T extends object>(target: T) => { deref(): T | undefined };

/**
 * The main class and control point of OpenSheetMusicDisplay.<br>
 * It can display MusicXML sheet music files in an HTML element container.<br>
 * After the constructor, use load() and render() to load and render a MusicXML file.
 */
interface OutsideMaskSegment {
    leftPx: number;
    topPx: number;
    widthPx: number;
    heightPx: number;
    color: string;
}

export class OpenSheetMusicDisplay {
    protected version: string = "2.3.1-dev"; // getter: this.Version
    // at release, bump version and change to -release, afterwards to -dev again
    private graphicNeedsPreparation: boolean = false;
    private lastMinMeasureToDrawIndex: number = 0;
    private lastMaxMeasureToDrawIndex: number = Number.MAX_SAFE_INTEGER;
    /** Native MusicXML fingerings temporarily replaced by setFingeringValues(). */
    private readonly nativeFingeringBackups: WeakMap<SourceStaffEntry, TechnicalInstruction[]> = new WeakMap();
    /** Staff entries with an active override. An empty override intentionally hides native fingerings. */
    private readonly fingeringOverrides: WeakSet<SourceStaffEntry> = new WeakSet();
    /**
     * Closest-to-note label position, retained while an override is empty. The staff skyline still
     * contains the removed native label, so querying it again would incorrectly push a replacement outward.
     */
    private readonly fingeringAnchors: WeakMap<SourceStaffEntry, number> = new WeakMap();
    private readonly inPlaceFingeringLabels: WeakSet<GraphicalLabel> = new WeakSet();

    /**
     * Creates and attaches an OpenSheetMusicDisplay object to an HTML element container.<br>
     * After the constructor, use load() and render() to load and render a MusicXML file.
     * @param container The container element OSMD will be rendered into.<br>
     *                  Either a string specifying the ID of an HTML container element,<br>
     *                  or a reference to the HTML element itself (e.g. div)
     * @param options An object for rendering options like the backend (svg/canvas) or autoResize.<br>
     *                For defaults see the OSMDOptionsStandard method in the [[OSMDOptions]] class.
     */
    constructor(container: string | HTMLElement,
                options: IOSMDOptions = OSMDOptions.OSMDOptionsStandard()) {
        // Store container element
        if (typeof container === "string") {
            // ID passed
            this.container = document.getElementById(<string>container);
        } else if (container && "appendChild" in <any>container) {
            // Element passed
            this.container = <HTMLElement>container;
        }
        if (!this.container) {
            throw new Error("Please pass a valid div container to OpenSheetMusicDisplay");
        }
        this.systemVirtualization = new SystemVirtualizationController(this.container);
        this.systemVirtualization.addLifecycleListener((event: ISystemLifecycleEvent): void => {
            if (event.type === "attached") {
                for (const root of event.roots) {
                    VexFlowGraphicalNote.flushDetachedSVGState(root);
                }
            }
        });

        // a copy: the object given isn't changed (it can be frozen, or used for another instance). Also for null as options.
        const optionsWithDefaults: IOSMDOptions = { ...options };
        if (optionsWithDefaults.autoResize === undefined) {
            optionsWithDefaults.autoResize = true;
        }
        this.backendType = BackendType.SVG; // default, can be changed by options
        this.setOptions(optionsWithDefaults);
    }

    /** Options from which OSMD creates cursors in enableOrDisableCursors(). */
    public cursorsOptions: CursorOptions[]; // set in the constructor (setOptions())
    public cursors: Cursor[] = [];
    public get cursor(): Cursor { // lowercase for backwards compatibility since cursor -> cursors change
        return this.cursors[0];
    }
    public get Cursor(): Cursor {
        return this.cursor;
    }
    /** The scale of the drawing, 1.0 by default: see Zoom, which sets it. */
    public zoom: number = 1.0;
    protected zoomUpdated: boolean = false;
    /** Timeout in milliseconds used in osmd.load(string) when string is a URL. */
    public loadUrlTimeout: number = 5000;

    protected container: HTMLElement;
    private readonly systemVirtualization: SystemVirtualizationController;
    private virtualizedInitialSystemCount: number | undefined;
    protected backendType: BackendType;
    protected needBackendUpdate: boolean;
    protected sheet: MusicSheet;
    protected drawer: VexFlowMusicSheetDrawer;
    protected drawBoundingBox: string;
    protected drawSkyLine: boolean;
    protected drawBottomLine: boolean;
    protected graphic: GraphicalMusicSheet;
    protected drawingParameters: DrawingParameters;
    protected rules: EngravingRules;
    private staffOpacityOverrides: Map<number, number> = new Map();
    private readAheadStaffEntryIndexGeneration: number = 0;
    private readonly readAheadStaffEntriesByMeasure: Map<number, GraphicalStaffEntry[]> = new Map();
    private readonly readAheadUnsortedMeasures: Set<number> = new Set();
    protected autoResizeEnabled: boolean;
    protected resizeHandlerAttached: boolean;
    /** The container's content width (getContainerContentWidth()) the last layout was made for: by render(), or at the start
     *  of incremental rendering. Turning autoResize on re-renders a rendered sheet only if the width differs from it. */
    protected layoutContainerWidth: number;
    /** Removes the container's ResizeObserver or the window's resize listener added for autoResize (see handleResize()),
     *  e.g. in dispose(). */
    private disposeResizeListener: () => void;
    /** The container's content width with the pages of the last render() in it, if they changed it: a sheet taller than
     *  the window makes the page's scrollbar appear (where scrollbars take space), which narrows a container as wide as the
     *  page. The container's observer (autoResize) doesn't render again for this width, which the render itself caused:
     *  the sheet stays laid out for the width without the scrollbar, as without autoResize, until the width changes for
     *  another reason or the app renders again. Another layout for it would cost about a third of a render (measured for
     *  PR 1992), and could change the width back, without end. */
    private containerWidthWithPages: number;
    /** Whether this instance has removed the pages that other instances left in the container (see removeLeftoverPages()). */
    private leftoverPagesRemoved: boolean = false;
    /** Whether destroy() was called: the instance can't load or render anymore. */
    private destroyed: boolean = false;
    /** The callbacks of the window's resize listener (see handleResize()), which reaches them only through this instance. */
    private resizeStartCallback: () => void;
    private resizeEndCallback: () => void;
    /** The key of the container's property that keeps this instance in memory while autoResize is on (see handleResize()). */
    private static readonly containerInstanceKey: symbol = Symbol("OpenSheetMusicDisplay");
    protected followCursor: boolean;
    /** Whether this instance uses the page's warm keeper, or "auto" (in Chromium-based browsers), see KeepWarm. */
    private keepWarm: boolean | "auto" = "auto";
    /** Whether this instance is counted in warmKeeperUsers (it rendered with KeepWarm on, and isn't disposed). */
    private usesWarmKeeper: boolean = false;
    /** The page's warm keeper (see KeepWarm): one instance, shared by all instances that use it. Static on purpose: its
     *  rendering keeps the compiled code alive for every instance, so one per page is enough. It holds no other instance. */
    private static warmKeeper: OpenSheetMusicDisplay = undefined;
    /** Whether the page's warm keeper was made or is scheduled, so that it's made once. */
    private static warmKeeperScheduled: boolean = false;
    /** Counts the times the page's warm keeper was freed: a keeper scheduled before that isn't made (or kept) anymore, also
     *  when another instance uses one again meanwhile, which schedules its own. */
    private static warmKeeperFreeCount: number = 0;
    /** How many instances use the warm keeper (see usesWarmKeeper); the last one to stop frees it. A count, not references,
     *  so that it keeps no instance in memory. */
    private static warmKeeperUsers: number = 0;
    /** The layout width in pixels of the warm keeper, which has no container in the page: getContainerContentWidth() returns it. */
    private fixedLayoutWidth: number = undefined;
    /** A function that is executed when the XML has been read.
     * The return value will be used as the actual XML OSMD parses,
     * so you can make modifications to the xml that OSMD will use.
     * By default it returns the XML unchanged. It can also be set by the onXMLRead option,
     * and osmd.setOptions() keeps it when the option is left out. */
    public OnXMLRead: (xml: string) => string = (xml: string): string => xml;
    public rangeSelection: ResolvedRangeSelectionConfig = createResolvedRangeSelectionConfig();
    private rangeInteractionOverlay: HTMLDivElement;
    private outsideMaskLayer: HTMLDivElement;
    private rangeChromeLayer: HTMLDivElement;
    private readonly outsideMaskPool: HTMLDivElement[] = [];
    private readonly dragHandleLines: [HTMLDivElement, HTMLDivElement] = [undefined, undefined];
    private maskDragChromePrepared: boolean = false;
    private rangeInteractionBoundElements: HTMLElement[] = [];
    private isRangeDragging: boolean = false;
    private hasActiveRangeSelectionOpacity: boolean = false;
    private rangeOpacityUpdateTimeoutId: number = 0;
    private lastRangeOpacityUpdateTimestampMs: number = 0;
    /** Cache key (selection + visible systems + opacity) of the last applied gray-out, so scroll-driven
     *  viewport updates can skip the expensive opacity recompute when nothing relevant changed. */
    private lastRangeOpacityKey: string = "";
    /** Whether the last gray-out pass also grayed decorations (clefs/tuplets/expressions), tracked so a
     *  note-only scroll pass doesn't skip a later decoration pass for the same viewport. */
    private lastRangeOpacityDecorationsApplied: boolean = false;
    private pendingRangePointerMoveAnchor: RangeSelectionAnchor;
    private rangePointerMoveAnimationFrameId: number = 0;
    private rangeTouchAutoScrollAnimationFrameId: number = 0;
    private rangeViewportUpdateAnimationFrameId: number = 0;
    private rangeViewportSettleUpdateTimeoutId: number = 0;
    private rangeViewportScrollTarget: HTMLElement | Window;
    private readonly rangeOpacityTouchedGraphicalNotes: Set<GraphicalNote> = new Set<GraphicalNote>();
    private readonly rangeOpacityTouchedNoteheadElements: Set<SVGElement> = new Set<SVGElement>();
    private readonly rangeOpacityTouchedElements: Set<Element> = new Set<Element>();
    /** Notes whose opacity was lowered by read-ahead (kept separate from range selection so they never clash). */
    private readonly readAheadOpacityTouchedGraphicalNotes: Set<GraphicalNote> = new Set<GraphicalNote>();
    private readonly readAheadOpacityTouchedLabels: Set<GraphicalLabel> = new Set<GraphicalLabel>();
    private readonly rangeSelectionElementCollector: RangeSelectionElementCollector = new RangeSelectionElementCollector();
    private needsCommittedRangeAnchorRefresh: boolean = false;
    private hoverAnchor: RangeSelectionAnchor;
    private dragStartAnchor: RangeSelectionAnchor;
    private dragCurrentAnchor: RangeSelectionAnchor;
    private pendingTouchRangeStartAnchor: RangeSelectionAnchor;
    private activeDragBound: "start" | "end" | "both" = "both";
    private rangeDragPointerCaptureElement: Element;
    private rangeDragPointerId: number = -1;
    private activeTouchPointerId: number = -1;
    private activeTouchStartClientX: number = 0;
    private activeTouchStartClientY: number = 0;
    private activeTouchMoved: boolean = false;
    private activeTouchDownAnchor: RangeSelectionAnchor;
    private activeTouchDragClientX: number = 0;
    private activeTouchDragClientY: number = 0;
    private touchDragScrollLockEnabled: boolean = false;
    private touchDragNativeScrollSuppressed: boolean = false;
    private isRangeHandleDragging: boolean = false;
    private touchPendingAction: "none" | "setOrCommit" | "clearSelection" = "none";
    private readonly rangePointerMoveListener: (event: PointerEvent) => void = (event: PointerEvent): void => this.onRangePointerMove(event);
    private readonly rangePointerDownListener: (event: PointerEvent) => void = (event: PointerEvent): void => this.onRangePointerDown(event);
    private readonly rangePointerUpListener: (event: PointerEvent) => void = (event: PointerEvent): void => this.onRangePointerUp(event);
    private readonly rangePointerCancelListener: (event: PointerEvent) => void = (event: PointerEvent): void => this.onRangePointerCancel(event);
    private readonly rangePointerLeaveListener: (event: PointerEvent) => void = (event: PointerEvent): void => this.onRangePointerLeave(event);
    private readonly touchMoveDuringRangeDragListener: (event: TouchEvent) => void = (event: TouchEvent): void => this.onTouchMoveDuringRangeDrag(event);
    private readonly rangeViewportUpdateListener: () => void = (): void => {
        this.scheduleRangeViewportUpdate();
        this.scheduleRangeViewportSettleUpdate();
    };
    /** A function that is called after each render: at the end of render() and of each batch of renderNext(), also of
     *  the renders OSMD starts itself (autoResize, enableIncrementalRenderingOnScroll()). A render replaces the SVG
     *  elements, so this is the place to color them or add listeners to them again. Also the option onRender. */
    public OnRender: (info: IRenderInfo) => void = undefined;
    /** Counts the calls of load(), clear() and dispose(). A load() whose number isn't the current one anymore was superseded:
     *  it doesn't set the sheet, and its promise is rejected (see load()). */
    private loadGeneration: number = 0;
    /** Aborts the URL request of the load() in progress, if any (see cancelPendingLoad()). */
    private loadAbortController: AbortController;

    /**
     * Load a MusicXML file
     * @param content is either the url of a file, or the root node of a MusicXML document,
     *   or the string content of a .xml/.mxl file, or a file blob.
     *   A string that starts with "<" (after whitespace) is read as XML, also without the XML declaration (<?xml ...?>).
     *   Any other string given here is taken as a URL and requested. That is only done for the string the app passes:
     *   what a Blob, an MXL file or the requested URL contains is never requested, it has to be MusicXML (or MXL).
     * @param tempTitle is used as the title for the piece if there is no title in the XML.
     *   The name or path of a MusicXML file (e.g. "scores/Sonata No. 1.musicxml") is used without its folder and extension.
     *   (The options, the EngravingRules and the zoom, osmd.Zoom, stay as they are set when a score is loaded.)
     * @returns a promise that is rejected with an Error when the score can't be loaded, for every kind of content:
     *   e.g. for malformed XML (the message has the XML parser's line and column), for a file that isn't MusicXML or MXL,
     *   or when reading the score fails (the Error's `cause` is then the reader's error). load() itself doesn't throw.
     *   A URL, a Blob and an MXL string are read asynchronously. When load(), clear() or dispose() is called before that
     *   has finished, the earlier load() doesn't set the sheet anymore, whichever file arrives first: its request is aborted,
     *   and its promise is rejected with an Error whose `name` is "AbortError" (like an aborted fetch()), so that a
     *   `.then(() => osmd.render())` of it doesn't run.
     * @example
     * osmd.load(url).then(() => osmd.render()).catch((error) => {
     *     if (error.name !== "AbortError") { // not an error: another score was loaded meanwhile
     *         showError(error.message);
     *     }
     * });
     */
    public load(content: string | Document | Blob, tempTitle: string = "Untitled Score"): Promise<{}> {
        if (this.destroyed) {
            return Promise.reject(this.createDestroyedError("load"));
        }
        this.cancelPendingLoad();
        const generation: number = this.loadGeneration;
        if (typeof AbortController === "function") {
            this.loadAbortController = new AbortController();
        }
        // the app's own string, so it can be a URL
        return this.continueLoad(content, { tempTitle, generation, isFileContent: false }).catch((error: Error) => {
            // e.g. the aborted request's error
            throw generation === this.loadGeneration ? error : this.createSupersededLoadError();
        });
    }

    /**
     * Supersedes a load() that hasn't finished: aborts its URL request, and makes its remaining steps reject instead of
     * setting the sheet (see load()).
     */
    private cancelPendingLoad(): void {
        this.loadGeneration++;
        this.loadAbortController?.abort();
        this.loadAbortController = undefined;
    }

    /** The Error for a call of the given method (e.g. "render") after destroy(). */
    private createDestroyedError(method: string): Error {
        return new Error(`OpenSheetMusicDisplay: ${method}() was called after destroy(). Create a new instance to show a score again.`);
    }

    /** The Error a superseded load() is rejected with (see load()). */
    private createSupersededLoadError(): Error {
        const error: Error = new Error(
            "OpenSheetMusicDisplay: load() was aborted, because load(), clear() or dispose() was called before it finished");
        error.name = "AbortError";
        return error;
    }

    /**
     * A step of load(): the first one, and each one that follows an asynchronous one (a request, unzipping, reading a Blob).
     * @param loadState what the steps of one load() pass along, with `isFileContent` set for the given content
     * @returns a promise that is rejected when the load() was superseded meanwhile, or when loadContent() throws or rejects
     */
    private continueLoad(content: string | Document | Blob, loadState: ILoadState): Promise<{}> {
        if (loadState.generation !== this.loadGeneration) {
            return Promise.reject(this.createSupersededLoadError());
        }
        try {
            return this.loadContent(content, loadState);
        } catch (error) {
            // string and Document content is read synchronously: without this, its errors would be thrown instead of rejected
            return Promise.reject(error);
        }
    }

    /** Does the work of a step of load() (see continueLoad()). Can throw as well as reject: continueLoad() turns both into
     *  a rejected promise.
     *  Only the string the app gave to load() can be a URL. A string that was read from a file (`loadState.isFileContent`: the text
     *  of a Blob, of an MXL archive's score file or of a URL's response) has to be XML or MXL and is rejected otherwise:
     *  if it were requested as a URL too, a file that a user picks could make the page send a request of the file's
     *  choice to the app's own server (with the user's cookies), and a URL whose response is a URL would be followed. */
    private loadContent(content: string | Document | Blob, loadState: ILoadState): Promise<{}> {
        this.reset();
        const self: OpenSheetMusicDisplay = this;
        /** For the steps that continue with what a file contains (a Blob, an MXL archive, a URL's response). */
        const fileContentState: ILoadState = { ...loadState, isFileContent: true };
        /** The XML string that `content` was parsed from, if it was given as a string. */
        let parsedXml: string;
        if (content instanceof Blob) {
            const mxlFile: MXLFile = new MXLFile(content);
            // check if this is a zip / mxl file
            return mxlFile.tryUnzip().then(() => {
                if (mxlFile.unzipSuccessful) {
                    return mxlFile.getXmlString().then(
                        (xmlString) => {
                            return self.continueLoad(xmlString, fileContentState);
                        },
                        (err: Error) => {
                            throw this.createError("OpenSheetMusicDisplay: Invalid MXL file", err);
                        }
                    );
                } else {
                    // not a zip
                    if (content instanceof Blob) { // always true. unfortunately need to check again for linter
                        return content.text().then((blobString) => {
                            return self.continueLoad(blobString, fileContentState);
                        });
                    }
                }
            });
        } else if (typeof content === "string") {
            const str: string = <string>content;
            // console.log("substring: " + str.substr(0, 5));
            if (str.startsWith("\x50\x4b\x03\x04")) {
                log.debug("[OSMD] This is a zip file, unpack it first: " + str);
                // This is a zip file, unpack it first
                return MXLHelper.MXLtoXMLstring(str).then(
                    (x: string) => {
                        return self.continueLoad(x, fileContentState);
                    },
                    (err: Error) => {
                        log.debug(err);
                        throw this.createError("OpenSheetMusicDisplay: Invalid MXL file", err);
                    }
                );
            }
            // Javascript loads strings as utf-16, which is wonderful BS if you want to parse UTF-8 :S
            if (str.startsWith("\uf7ef\uf7bb\uf7bf")) {
                log.debug("[OSMD] UTF with BOM detected, truncate first 3 bytes and pass along: " + str);
                // UTF with BOM detected, truncate first three bytes and pass along
                return self.continueLoad(str.substring(3), loadState); // (the same content)
            }
            let trimmedStr: string = str;
            if (/^\s/.test(trimmedStr)) { // only trim if we need to. (end of string is irrelevant)
                trimmedStr = trimmedStr.trim(); // trim away empty lines at beginning etc., also a byte order mark (U+FEFF)
            }
            // XML, with or without the XML declaration (<?xml ...?>), which is optional: e.g. a MusicXML string from jsdom's
            //   XMLSerializer starts with <!DOCTYPE, one from Python's ElementTree with <score-partwise>. A URL can't start with "<".
            if (trimmedStr.startsWith("<")) {
                const modifiedXml: string = this.OnXMLRead(trimmedStr); // by default just returns trimmedStr unless a function options.OnXMLRead was set.
                log.debug("[OSMD] Finally parsing XML content, length: " + modifiedXml.length);
                // Parse the string representing an xml file
                const parser: DOMParser = new DOMParser();
                // (without the white space between the elements: fewer nodes to parse, see removeWhitespaceBetweenElements())
                content = parser.parseFromString(removeWhitespaceBetweenElements(modifiedXml), "application/xml");
                parsedXml = modifiedXml;
            } else if (loadState.isFileContent) {
                return Promise.reject(new Error("OpenSheetMusicDisplay: The file is not MusicXML: " + (trimmedStr.length === 0 ?
                    "it is empty" : `it starts with ${JSON.stringify(trimmedStr.substring(0, 20))} instead of "<"`)));
            } else if (trimmedStr.length < 2083) {
                // The app's string, and not XML: a URL. There is no format to check: a relative URL can be nearly any
                //   text ("score.musicxml", "scores/My Song (2).mxl", "?id=5"). Only the length tells content from a URL:
                //   2083 characters were the limit of Internet Explorer's URLs.
                log.debug("[OSMD] Retrieve the file at the given URL: " + trimmedStr);
                return AJAX.ajax(trimmedStr, this.loadUrlTimeout, this.loadAbortController?.signal).then(
                    (s: string) => { return self.continueLoad(s, fileContentState); },
                    (exc: Error) => { throw exc; }
                );
            } else {
                return Promise.reject(new Error("OpenSheetMusicDisplay: The string is neither XML nor a URL: it doesn't start with \"<\", " +
                    `and is too long for a URL (${trimmedStr.length} characters)`));
            }
        }

        if (!content || !(<any>content).nodeName) {
            return Promise.reject(new Error("OpenSheetMusicDisplay: The document which was provided is invalid"));
        }
        const xmlDocument: Document = (<Document>content);
        if (typeof xmlDocument.getElementsByTagName === "function") { // (not for a node given instead of a document)
            let parserError: string = this.getXmlParserError(xmlDocument);
            if (parserError) {
                if (parsedXml) {
                    // parse the string as it was given (with its line breaks), so that the error's line and column are the file's
                    parserError = this.getXmlParserError(new DOMParser().parseFromString(parsedXml, "application/xml")) ?? parserError;
                }
                return Promise.reject(new Error("OpenSheetMusicDisplay: The XML is malformed: " + parserError));
            }
        }
        const xmlDocumentNodes: NodeList = xmlDocument.childNodes;
        log.debug("[OSMD] load(), Document url: " + xmlDocument.URL);

        let scorePartwiseElement: Element;
        for (let i: number = 0, length: number = xmlDocumentNodes.length; i < length; i += 1) {
            const node: Node = xmlDocumentNodes[i];
            if (node.nodeType === Node.ELEMENT_NODE && node.nodeName.toLowerCase() === "score-partwise") {
                scorePartwiseElement = <Element>node;
                break;
            }
        }
        if (!scorePartwiseElement) {
            console.error("Could not parse MusicXML, no valid partwise element found");
            return Promise.reject(new Error("OpenSheetMusicDisplay: Document is not a valid 'partwise' MusicXML"));
        }
        const score: IXmlElement = new IXmlElement(scorePartwiseElement);
        const temposCalculator: TemposCalculator = new TemposCalculator();
        const reader: MusicSheetReader = new MusicSheetReader([temposCalculator], this.rules);
        this.sheet = reader.createMusicSheet(score, loadState.tempTitle);
        if (this.sheet === undefined) {
            // error loading sheet, already logged by the reader
            return Promise.reject(this.createError(
                "OpenSheetMusicDisplay: The given music sheet was incomplete or could not be loaded", reader.ReadingError));
        }
        // if (this.sheet.TitleString === "osmd.Version") {
        //     this.sheet.TitleString = "OSMD version: " + this.Version; // useful for debug e.g. when console not available
        // }
        log.info(`[OSMD] Loaded sheet ${this.sheet.TitleString} successfully.`);

        this.needBackendUpdate = true;
        try {
            this.updateGraphic();
        } catch (error) {
            this.reset(); // no sheet without a graphic: after a failed load(), the instance has neither, as before its first load()
            throw error;
        }

        return Promise.resolve({});
    }

    /** Creates an Error with the given message, followed by the message of the error that caused it, which is kept as `cause`. */
    private createError(message: string, cause: Error): Error {
        const causeMessage: string = cause?.message ?? (cause ? String(cause) : undefined);
        const error: Error & { cause?: unknown } = new Error(causeMessage ? `${message}: ${causeMessage}` : message);
        if (cause) {
            error.cause = cause; // (the constructor's cause option is ES2022)
        }
        return error;
    }

    /**
     * Returns the XML parser's message if the document isn't well-formed XML, e.g. "error on line 3 at column 28:
     * Unexpected token inside entity" (Chrome) or "XML Parsing Error: not well-formed Line Number 3, Column 28" (Firefox),
     * and undefined for a well-formed document.
     * DOMParser doesn't throw for malformed XML, it returns a document with a <parsererror> element: Firefox as the root
     * element, Chrome and Safari inside of the document as far as it was read (so the root can still be <score-partwise>,
     * with the rest of the score missing).
     */
    private getXmlParserError(xmlDocument: Document): string {
        const errorElement: Element = xmlDocument.getElementsByTagName("parsererror")[0];
        if (!errorElement) {
            return undefined;
        }
        // Chrome, Safari: <h3>This page contains the following errors:</h3><div>the message</div><h3>Below is a rendering ...</h3>
        // Firefox: the message as text ("XML Parsing Error: ...", "Location: <the page's URL>", "Line Number ..., Column ...:"),
        //   then <sourcetext> with the line of the error, which can be the whole file
        const messageNode: Node = errorElement.getElementsByTagName("div")[0] ?? errorElement.firstChild ?? errorElement;
        const lines: string[] = (messageNode.textContent ?? "").split("\n")
            .map((line: string): string => line.trim())
            .filter((line: string): boolean => line.length > 0 && !line.startsWith("Location:"));
        const message: string = lines.join(" ").replace(/:$/, "");
        return message.length > 0 ? message.substring(0, 300) : "unknown error";
    }

    /**
     * (Re-)creates the graphic sheet from the music sheet, and puts the cursors back to the start.
     * Needed e.g. after changing a rule that applies when the graphic is created, like EngravingRules.RenderHiddenNotes.
     * A changed transposition (osmd.Sheet.Transpose, Instrument.Transpose, osmd.TransposeCalculator) or a voice shown or
     * hidden (Voice.Visible, Instrument.Visible) doesn't need it: render() creates the graphic again then, and keeps the
     * cursors where they are.
     */
    public updateGraphic(): void {
        // the draw range of the sheet into the rules, as render() does: the cursors start at its first measure. The rules can still have
        //   the indices of the sheet before (e.g. index 2 for drawFromMeasureNumber 3 without a pickup measure: measure 2 with one).
        [this.rules.MinMeasureToDrawIndex, this.rules.MaxMeasureToDrawIndex] = this.getDrawRangeIndices();
        this.createGraphic();
        if (this.drawingParameters.drawCursors) {
            this.cursors.forEach(cursor => {
                cursor.init(this.sheet.MusicPartManager, this.graphic);
            });
        }
    }

    /** Creates the graphic sheet from the music sheet: its graphical measures, notes and chord symbols, with the notes and chord
     *  symbols transposed by the current transposition (see GraphicalMusicSheet.TranspositionChanged). Unlike updateGraphic(), it
     *  leaves the cursors alone: render() and renderNext() create them again on the new graphic, where they were.
     *  @param drawRange the first and the last measure to draw (source measure indices, see getDrawRangeIndices()): the multiple
     *   rest measures begin and end with them (see MusicSheetCalculator.prepareGraphicalMusicSheet())
     */
    private createGraphic(drawRange: [number, number] = this.getDrawRangeIndices()): void {
        const calc: MusicSheetCalculator = new VexFlowMusicSheetCalculator(this.rules);
        // the rules have the draw range while the graphic is created, then their own values again (e.g. a draw range set during
        //   incremental rendering, for the next session)
        const rulesDrawRange: [number, number] = [this.rules.MinMeasureToDrawIndex, this.rules.MaxMeasureToDrawIndex];
        [this.rules.MinMeasureToDrawIndex, this.rules.MaxMeasureToDrawIndex] = drawRange;
        try {
            this.graphic = new GraphicalMusicSheet(this.sheet, calc, false);
        } finally {
            [this.rules.MinMeasureToDrawIndex, this.rules.MaxMeasureToDrawIndex] = rulesDrawRange;
        }
        this.graphicNeedsPreparation = true;
        if (this.drawingParameters.DrawingParametersEnum === DrawingParametersEnum.leadsheet) {
            this.graphic.LeadSheet = true;
        }
    }

    /** Whether the graphic has to be created again before it's rendered: after a change of the transposition or of the
     *  visible voices, which apply when it's created (see GraphicalMusicSheet.TranspositionChanged and VoiceVisibilityChanged),
     *  or for a draw range that would divide the multiple rest measures otherwise (see MusicSheetCalculator.multiRestsFitDrawRange()).
     *  @param drawRange the first and the last measure to draw (source measure indices, see getDrawRangeIndices())
     */
    private graphicOutdated(drawRange: [number, number]): boolean {
        return this.graphic.TranspositionChanged || this.graphic.VoiceVisibilityChanged ||
            !this.graphic.GetCalculator.multiRestsFitDrawRange(drawRange[0], drawRange[1]);
    }

    /** Lazy rendering (LazyConsistentGraphic): number of systems already drawn into the shared
     *  backends across prior batches, counted through the pages. Greedy layout is *usually* forward-stable, so the next
     *  batch skips redrawing these and draws from this index -- but some scores re-position earlier systems as the
     *  prefix grows, so each batch verifies the drawn systems against lazyDrawnSystemPositions and redraws from
     *  the topmost one that moved (reconciliation). */
    private lazyDrawnSystemCount: number = 0;
    /** Lazy rendering: the page and the absolute Y (in units, on the page) each already-drawn system [index] was drawn at,
     *  used to detect when a later batch's full-prefix layout moves an earlier system (forward-stability
     *  is not universal) so it can be redrawn at its corrected position. */
    private lazyDrawnSystemPositions: { pageNumber: number, y: number }[] = [];
    /** Lazy HORIZONTAL rendering (RenderSingleHorizontalStaffline): number of graphical measures of the top
     *  staffline already drawn, counted through the systems in order (see renderAppendGrowingHorizontal()).
     *  The next batch draws from here. */
    private lazyDrawnHMeasureCount: number = 0;

    /** Incremental rendering ({@link renderNext}): whether a session is in progress (started, not yet reset). */
    private lazyIncrementalActive: boolean = false;
    /** Incremental rendering: source-measure indices of the first and the last measure the session renders, the ones
     *  render() draws (the draw range, see getDrawRangeIndices()), limited to the sheet. Set when the session starts. */
    private lazyFirstMeasureIndex: number = 0;
    private lazyLastMeasureIndex: number = 0;
    /** Incremental rendering: source-measure index of the last measure the session draws, for its progress: lazyLastMeasureIndex,
     *  or the last measure of the last system drawn once a batch's layout ends there before it (drawUpToSystemNumber,
     *  drawUpToPageNumber, see lazyLayoutEndBefore()). Set by the session's first batch. */
    private lazyLayoutLastMeasureIndex: number = 0;
    /** Incremental rendering: source-measure index where the next batch continues (the drawn frontier). */
    private lazyNextSourceIndex: number = 0;
    /** Incremental rendering of the endless page: the width in pixels all batches of the session lay the sheet out at, the
     *  container's content width when the first batch read it (like render() reads it once). A batch reading the width again
     *  laid its systems out narrower than the ones drawn before when the container had become narrower in between, e.g. by
     *  the vertical scrollbar the page gets once the first batch makes it longer than the window. */
    private lazyLayoutWidth: number = 0;
    /** Incremental rendering: the scroll listener + its target, while enableIncrementalRenderingOnScroll() is on. */
    private lazyScrollHandler: (() => void) | undefined;
    private lazyScrollTarget: HTMLElement | Window | undefined;

    /** Whether the container is observed for autoResize (see handleResize()), instead of the window. */
    private observesContainer(): boolean {
        return this.resizeHandlerAttached && typeof ResizeObserver === "function";
    }

    /**
     * Whether a render is skipped because the container has no width, e.g. a hidden one (display: none, a tab that isn't
     * shown) or one that isn't in the document: a layout for width 0 puts one measure into each system. Logs a warning then.
     * With autoResize, the container's observer renders the sheet (render()) once the container has a width.
     * A single horizontal staffline doesn't depend on the width: it's rendered.
     * @param method the method that renders, for the warning: "render" or "renderNext"
     */
    private skipsRenderWithoutWidth(method: string): boolean {
        if (this.getContainerContentWidth() !== 0 || this.rules.RenderSingleHorizontalStaffline) {
            return false;
        }
        this.layoutContainerWidth = 0; // the width of the layout that is due: the observer renders at another one
        log.warn(`[OSMD] ${method}(): the container's width is 0 (e.g. it's hidden, or not in the document), so the sheet wasn't rendered. ` +
            (this.observesContainer() ? "It will be rendered when the container gets a width (autoResize)."
                : "Call render() again when the container has a width."));
        return true;
    }

    /** Render the loaded music sheet to the container.
     *  A changed transposition (osmd.Sheet.Transpose, Instrument.Transpose, osmd.TransposeCalculator) or a voice shown or hidden
     *  (Voice.Visible, Instrument.Visible) is applied without updateGraphic(): render() creates the graphic again then,
     *  see GraphicalMusicSheet.TranspositionChanged and VoiceVisibilityChanged. It does so too when the draw range
     *  (drawFromMeasureNumber, drawUpToMeasureNumber) begins or ends inside a multiple rest measure, which is divided there,
     *  or one was divided for the previous draw range (see MusicSheetCalculator.multiRestsFitDrawRange()).
     */
    public render(): void {
        if (this.destroyed) {
            throw this.createDestroyedError("render");
        }
        if (!this.graphic) {
            throw new Error("OSMD: load() needs to be called before render()");
        }
        if (this.skipsRenderWithoutWidth("render")) {
            return;
        }
        // drawing range, adjusted for a pickup measure
        const drawRange: [number, number] = this.getDrawRangeIndices();
        // The notes and chord symbols are transposed when the graphic is created, the key signatures in each layout (reCalculate()):
        //   after a change of the transposition, create the graphic again, otherwise only the key signatures would be transposed.
        //   The notes of hidden voices are left out when the graphic is created too, and the multiple rest measures begin and end
        //   with the draw range.
        //   Unlike updateGraphic(), this keeps the cursors where they are (enableOrDisableCursors() below creates them on the new graphic).
        if (this.graphicOutdated(drawRange)) {
            this.createGraphic(drawRange);
        }
        // A full render() supersedes any incremental render in progress: abandon it.
        this.resetIncrementalRendering();
        // A normal (non-lazy) render never uses the lazy reuse caches; force the gate off so a prior
        // lazy session can't make them affect this render (they aren't cleared for normal renders).
        this.rules.LazyConsistentGraphic = false;
        this.rangeSelectionElementCollector.invalidate();
        this.drawer?.clear(); // clear canvas before setting width
        // this.graphic.GetCalculator.clearSystemsAndMeasures(); // maybe?
        // this.graphic.GetCalculator.clearRecreatedObjects();

        [this.rules.MinMeasureToDrawIndex, this.rules.MaxMeasureToDrawIndex] = drawRange;

        // Set page width: the one the pages get as well (see createOrRefreshRenderBackend() below)
        let width: number = this.getContainerContentWidthWithoutPages();
        this.layoutContainerWidth = width;
        if (this.rules.RenderSingleHorizontalStaffline) {
            width = this.rules.SheetMaximumWidth; // set safe maximum (browser limit), will be reduced later
            // reduced later in MusicSheetCalculator.calculatePageLabels (sets sheet.pageWidth to page.PositionAndShape.Size.width before labels)
            // rough calculation:
            // width = 600 * this.sheet.SourceMeasures.length;
        }
        // log.debug("[OSMD] render width: " + width);

        this.sheet.pageWidth = width / this.zoom / 10.0;
        if (this.rules.PageFormat && !this.rules.PageFormat.IsUndefined) {
            this.rules.PageHeight = this.sheet.pageWidth / this.rules.PageFormat.aspectRatio;
            log.debug("[OSMD] PageHeight: " + this.rules.PageHeight);
        } else {
            log.debug("[OSMD] endless/undefined pageformat, id: " + this.rules.PageFormat.idString);
            this.rules.PageHeight = 100001; // infinite page height // TODO maybe Number.MAX_VALUE or Math.pow(10, 20)?
        }

        // Before introducing the following optimization (maybe irrelevant), tests
        // have to be modified to ensure that width is > 0 when executed
        //if (isNaN(width) || width === 0) {
        //    return;
        //}

        // Rebuild measures when drawing range changed so state like 8va spans is seeded correctly
        const currentMinIndex: number = this.rules.MinMeasureToDrawIndex;
        const currentMaxIndex: number = this.rules.MaxMeasureToDrawIndex;
        if (this.graphicNeedsPreparation ||
            this.lastMinMeasureToDrawIndex !== currentMinIndex || this.lastMaxMeasureToDrawIndex !== currentMaxIndex) {
            this.prepareGraphic();
        }

        // Calculate again
        this.graphic.reCalculate();

        if (this.drawingParameters.drawCursors) {
            this.graphic.Cursors.length = 0;
        }

        // needBackendUpdate is well intentioned, but we need to cover all cases.
        //   backends also need an update when this.zoom was set from outside, which unfortunately doesn't have a setter method to set this in.
        //   so just for compatibility, we need to assume users set osmd.zoom, so we'd need to check whether it was changed compared to last time.
        if (true || this.needBackendUpdate) {
            this.createOrRefreshRenderBackend(this.layoutContainerWidth); // pages as wide as the layout
            this.needBackendUpdate = false;
        }
        // The pages, sized for the layout, can change the container's width, e.g. when the page's scrollbar appears:
        //   the container's observer doesn't render again for that (see containerWidthWithPages).
        this.containerWidthWithPages = undefined;
        if (this.observesContainer()) {
            const widthWithPages: number = this.getContainerContentWidth();
            if (widthWithPages !== this.layoutContainerWidth) {
                this.containerWidthWithPages = widthWithPages;
            }
        }

        this.drawer.setZoom(this.zoom);

        for (const measure of this.sheet.SourceMeasures) {
            measure.WasRendered = false;
        }
        // Finally, draw. A virtualized first render keeps the complete graphical layout/data but only
        // materializes the first few systems; the controller draws later systems from this same layout.
        if (this.virtualizedInitialSystemCount !== undefined) {
            this.drawer.LazyDrawSystemsFromIndex = 0;
            this.drawer.LazyDrawSystemsToIndexExcl = this.virtualizedInitialSystemCount;
        }
        this.drawer.drawSheet(this.graphic);
        this.drawer.LazyDrawSystemsFromIndex = -1;
        this.drawer.LazyDrawSystemsToIndexExcl = Number.POSITIVE_INFINITY;

        if (this.virtualizedInitialSystemCount !== undefined) {
            // The full graphical score exists and is safe for iterator/cursor/timing consumers even where
            // SVG has not been materialized yet.
            for (const measure of this.sheet.SourceMeasures) {
                measure.WasRendered = true;
            }
            this.configureFullLayoutSystemVirtualization();
        }

        this.enableOrDisableCursors(this.drawingParameters.drawCursors);

        if (this.drawingParameters.drawCursors) {
            // Update the cursor position
            this.cursors.forEach(cursor => {
                cursor.update();
            });
        }
        this.reapplyStaffOpacityOverrides();
        this.syncInteractiveRangeSelection();
        this.needsCommittedRangeAnchorRefresh = true;
        // The SVG was rebuilt, so the previously applied gray-out attributes are gone. Invalidate the
        // opacity cache key so the upcoming renderRangeSelection() actually re-applies it.
        this.lastRangeOpacityKey = "";
        this.lastRangeOpacityDecorationsApplied = false;
        this.renderRangeSelection();
        this.systemVirtualization.refresh();
        this.zoomUpdated = false;
        this.rules.RenderCount++;
        //console.log("[OSMD] render finished");
        this.useWarmKeeper();
        if (this.OnRender) {
            this.OnRender({ incremental: false, done: true });
        }
    }

    /**
     * Whether this instance uses the page's warm keeper: true, false, or "auto" (default): in Chromium-based browsers (Chrome,
     * Edge, Opera, Electron, ...), see isChromiumBrowser(). The keeper is a small rendering OSMD keeps alive so that a sheet
     * rendered after clear() renders about as fast as the one before it. Also the option keepWarm (constructor, setOptions()).
     * The JavaScript engine (V8 in Chrome, Edge and Node) compiles OSMD's code for the objects it creates, and throws that
     * compiled code away once no such object is left, at a full garbage collection. After clear(), nothing of the sheet is
     * left, so the next sheet then rendered up to ~27% slower if such a collection happened in between (Chrome).
     * Firefox and Safari showed no such slowdown, and no gain from the keeper: "auto" makes none there.
     * After the first render() or renderNext() of an instance that uses it, OSMD renders a built-in score of 8 measures
     * (WarmKeeperScore.ts) when the page is idle (requestIdleCallback), into a container that isn't in the page: about
     * 20 ms in Chrome, and a few MB. One keeper per page, shared by the instances that use it, as any rendering keeps the
     * compiled code alive. It's freed when the last of them is disposed (dispose()) or turns this off. Not made without
     * requestIdleCallback (e.g. Node with jsdom).
     */
    public get KeepWarm(): boolean | "auto" {
        return this.keepWarm;
    }

    public set KeepWarm(value: boolean | "auto") {
        this.keepWarm = value;
        if (!this.keepsWarm()) {
            this.stopUsingWarmKeeper();
        }
    }

    /** Whether this instance uses the page's warm keeper with its KeepWarm setting in this browser ("auto" decided). */
    private keepsWarm(): boolean {
        return this.keepWarm === true || this.keepWarm === "auto" && typeof navigator !== "undefined" &&
            OpenSheetMusicDisplay.isChromiumBrowser(navigator);
    }

    /**
     * Whether `browser` (the page's navigator) is a Chromium-based browser, whose JavaScript engine V8 throws away the code it
     * compiled for objects of which none is left (see KeepWarm): from navigator.userAgentData (Chromium 90+, secure pages
     * only), else from the user agent: Chromium-based browsers name "Chrome/" or "Chromium/" in it, Safari and Firefox don't
     * (also not on iOS, where every browser is WebKit-based: CriOS, EdgiOS, FxiOS).
     */
    private static isChromiumBrowser(browser: Navigator): boolean {
        const brands: { brand: string }[] = (<any>browser).userAgentData?.brands;
        if (Array.isArray(brands) && brands.length > 0) {
            return brands.some((b: { brand: string }) => b.brand === "Chromium");
        }
        return /Chrome\/|Chromium\//.test(browser.userAgent ?? "");
    }

    /** Stops using the page's warm keeper (see KeepWarm): the last instance to stop frees it. */
    private stopUsingWarmKeeper(): void {
        if (!this.usesWarmKeeper) {
            return;
        }
        this.usesWarmKeeper = false;
        OpenSheetMusicDisplay.warmKeeperUsers--;
        if (OpenSheetMusicDisplay.warmKeeperUsers > 0) {
            return;
        }
        const keeper: OpenSheetMusicDisplay = OpenSheetMusicDisplay.warmKeeper;
        OpenSheetMusicDisplay.warmKeeper = undefined;
        OpenSheetMusicDisplay.warmKeeperScheduled = false;
        OpenSheetMusicDisplay.warmKeeperFreeCount++;
        keeper?.clear();
        keeper?.dispose();
    }

    /**
     * Uses the page's warm keeper (see KeepWarm), after a render() or renderNext(): makes it when the page is idle, if no
     * instance did yet. The keeper is a separate instance that renders the built-in score into a container that isn't in
     * the page.
     */
    private useWarmKeeper(): void {
        // Browsers only: without requestIdleCallback (e.g. Node with jsdom) there's no idle time to render the keeper in,
        //   and scripts that render many sheets (e.g. generateImages_browserless) would render a keeper for each instance.
        if (!this.keepsWarm() || this.fixedLayoutWidth !== undefined || typeof (<any>globalThis).requestIdleCallback !== "function" ||
            typeof document === "undefined" || typeof atob === "undefined") {
            return;
        }
        if (!this.usesWarmKeeper) {
            this.usesWarmKeeper = true;
            OpenSheetMusicDisplay.warmKeeperUsers++;
        }
        if (OpenSheetMusicDisplay.warmKeeperScheduled) {
            return; // made or scheduled already
        }
        OpenSheetMusicDisplay.warmKeeperScheduled = true;
        const freeCount: number = OpenSheetMusicDisplay.warmKeeperFreeCount;
        /** Whether the keeper scheduled here was freed meanwhile: no instance used it anymore. Checking only that some
         *  instance uses one wasn't enough: after it was freed, another instance's render() schedules its own. Both
         *  were made, and the later one replaced the other in warmKeeper, which then wasn't freed. */
        const freed: () => boolean = (): boolean => OpenSheetMusicDisplay.warmKeeperFreeCount !== freeCount;
        const makeKeeper: () => void = (): void => {
            if (freed()) {
                return;
            }
            const keeper: OpenSheetMusicDisplay = new OpenSheetMusicDisplay(document.createElement("div"),
                { autoResize: false, drawingParameters: "default", keepWarm: false });
            keeper.fixedLayoutWidth = 1200;
            keeper.load(atob(WarmKeeperScoreMxlBase64), "OSMD warm keeper").then(() => {
                if (freed()) {
                    keeper.clear();
                    return;
                }
                keeper.render();
                OpenSheetMusicDisplay.warmKeeper = keeper;
            }).catch((e: any) => {
                log.debug("[OSMD] warm keeper not rendered: " + e);
            });
        };
        (<any>globalThis).requestIdleCallback(makeKeeper, { timeout: 3000 }); // at the latest after 3 s
    }

    /**
     * The measures render() draws, as source-measure indices: those of the draw range set with setOptions() (drawFromMeasureNumber,
     * drawUpToMeasureNumber, kept in EngravingRules.MinMeasureToDrawNumber and MaxMeasureToDrawNumber), where measure number n is
     * at index n instead of n - 1 if the sheet begins with a pickup measure (number 0). Without a measure number set (the defaults
     * 0 and Number.MAX_VALUE), the index in EngravingRules.MinMeasureToDrawIndex or MaxMeasureToDrawIndex, e.g. one set there.
     * load(), render() and renderNext() put these indices into the rules: the ones there are those of the sheet loaded.
     * @returns the indices of the first and the last measure. The last one can be beyond the sheet (by default
     *  Number.MAX_VALUE), and if the first one is, the layout starts at the sheet's last measure.
     */
    private getDrawRangeIndices(): [number, number] {
        // The indices in the rules can be those of the sheet laid out before, adjusted to its pickup measure (and the first one to
        //   its measures, see MusicSheetCalculator.calculateMusicSystems()), so they're found again from the measure numbers set:
        //   e.g. with drawFromMeasureNumber 3, a sheet without a pickup measure after one with a pickup measure was drawn from measure 4.
        const pickup: boolean = !!this.Sheet.SourceMeasures[0]?.ImplicitMeasure;
        let firstIndex: number = this.rules.MinMeasureToDrawIndex;
        const fromNumber: number = this.rules.MinMeasureToDrawNumber;
        if (fromNumber > 0) {
            // from measure 1, the pickup measure is drawn too
            firstIndex = pickup && fromNumber > 1 ? fromNumber : Math.max(fromNumber - 1, 0);
        }
        let lastIndex: number = this.rules.MaxMeasureToDrawIndex;
        const upToNumber: number = this.rules.MaxMeasureToDrawNumber;
        if (upToNumber < Number.MAX_VALUE) {
            lastIndex = pickup && upToNumber > 0 ? upToNumber : Math.max(upToNumber - 1, 0);
        }
        return [firstIndex, lastIndex];
    }

    /**
     * Calculate the complete graphical score, but initially draw only a few systems. This preserves access
     * to every measure/note for timing and playback while avoiding full-score SVG creation at startup.
     * Call {@link enableSystemVirtualization} before or after this method to materialize viewport systems.
     */
    public renderVirtualized(options?: { initialSystems?: number }): void {
        if (this.backendType !== BackendType.SVG || this.rules.RenderSingleHorizontalStaffline) {
            this.render();
            return;
        }
        this.virtualizedInitialSystemCount = Math.max(1, Math.floor(options?.initialSystems ?? 3));
        try {
            this.render();
        } finally {
            this.virtualizedInitialSystemCount = undefined;
        }
    }

    private configureFullLayoutSystemVirtualization(): void {
        const descriptors: IVirtualSystemDescriptor[] = [];
        for (let pageIndex: number = 0; pageIndex < this.graphic.MusicPages.length; pageIndex++) {
            const page: GraphicalMusicPage = this.graphic.MusicPages[pageIndex];
            const renderRoot: HTMLElement = this.drawer.Backends[pageIndex]?.getRenderElement();
            const svg: SVGSVGElement = renderRoot?.querySelector<SVGSVGElement>("svg");
            if (!svg) {
                continue;
            }
            for (let systemIndex: number = 0; systemIndex < page.MusicSystems.length; systemIndex++) {
                const bounds: RectangleF2D = this.drawer.getSystemPixelBounds(page.MusicSystems[systemIndex]);
                descriptors.push({
                    key: `${page.PageNumber}:${systemIndex}`,
                    svg,
                    top: bounds.y,
                    bottom: bounds.y + bounds.height
                });
            }
        }
        this.systemVirtualization.configureExpectedSystems(descriptors, (keys: string[]): SVGGElement[] => {
            const indexesByPage: Map<number, number[]> = new Map<number, number[]>();
            for (const key of keys) {
                const [pageNumberString, systemIndexString] = key.split(":");
                const pageIndex: number = Number.parseInt(pageNumberString, 10) - 1;
                const systemIndex: number = Number.parseInt(systemIndexString, 10);
                if (!Number.isFinite(pageIndex) || !Number.isFinite(systemIndex)) {
                    continue;
                }
                const indexes: number[] = indexesByPage.get(pageIndex) ?? [];
                indexes.push(systemIndex);
                indexesByPage.set(pageIndex, indexes);
            }
            const firstNewGroup: number = this.drawer.SystemGroups.length;
            for (const [pageIndex, indexes] of indexesByPage) {
                this.drawer.drawSystemIndexes(this.graphic, pageIndex, indexes.sort((a, b): number => a - b));
            }
            this.rangeSelectionElementCollector.invalidate();
            this.reapplyStaffOpacityOverrides(this.drawer.SystemGroups.slice(firstNewGroup));
            this.lastRangeOpacityKey = "";
            this.lastRangeOpacityDecorationsApplied = false;
            this.renderRangeSelection();
            if (this.drawingParameters.drawCursors) {
                this.cursors.forEach(cursor => cursor.update());
            }
            return this.drawer.SystemGroups.slice(firstNewGroup);
        }, (key: string, root: SVGGElement): void => {
            this.releaseVirtualizedSystem(key, root);
        });
    }

    private releaseVirtualizedSystem(key: string, root: SVGGElement): void {
        const [pageNumberString, systemIndexString] = key.split(":");
        const pageIndex: number = Number.parseInt(pageNumberString, 10) - 1;
        const systemIndex: number = Number.parseInt(systemIndexString, 10);
        const system: MusicSystem = this.graphic?.MusicPages[pageIndex]?.MusicSystems[systemIndex];
        if (system) {
            for (const staffLine of system.StaffLines) {
                for (const measure of staffLine.Measures) {
                    for (const staffEntry of measure?.staffEntries ?? []) {
                        for (const voiceEntry of staffEntry.graphicalVoiceEntries ?? []) {
                            for (const note of voiceEntry.notes ?? []) {
                                note.releaseRenderedSVG();
                            }
                        }
                        for (const label of staffEntry.FingeringEntries ?? []) {
                            if (label.SVGNode && root.contains(label.SVGNode)) {
                                label.SVGNode = undefined;
                            }
                        }
                        for (const lyric of staffEntry.LyricsEntries ?? []) {
                            const label: GraphicalLabel = lyric.GraphicalLabel;
                            if (label?.SVGNode && root.contains(label.SVGNode)) {
                                label.SVGNode = undefined;
                            }
                        }
                        for (const chord of staffEntry.graphicalChordContainers ?? []) {
                            const label: GraphicalLabel = chord.GraphicalLabel;
                            if (label?.SVGNode && root.contains(label.SVGNode)) {
                                label.SVGNode = undefined;
                            }
                        }
                    }
                }
                for (const slur of staffLine.GraphicalSlurs ?? []) {
                    if (slur.SVGElement && root.contains(slur.SVGElement)) {
                        slur.SVGElement = undefined;
                    }
                }
            }
            const pending: BoundingBox[] = [system.PositionAndShape];
            while (pending.length > 0) {
                const box: BoundingBox = pending.pop();
                const data: { SVGNode?: Node, SVGElement?: Node } = box?.DataObject as { SVGNode?: Node, SVGElement?: Node };
                if (data?.SVGNode && root.contains(data.SVGNode)) {
                    data.SVGNode = undefined;
                }
                if (data?.SVGElement && root.contains(data.SVGElement)) {
                    data.SVGElement = undefined;
                }
                for (const child of box?.ChildElements ?? []) {
                    pending.push(child);
                }
            }
        }
        this.drawer.forgetSystemGroup(root);
        this.rangeSelectionElementCollector.invalidate();
        for (const label of this.readAheadOpacityTouchedLabels) {
            if (label.SVGNode && root.contains(label.SVGNode)) {
                label.SVGNode = undefined;
            }
        }
    }

    /** Internal range-based engine behind {@link renderNext} (the public incremental API). Lays out the
     *  prefix of the draw range up to toMeasureIndex and APPENDS the newly-stable source measures below previously
     *  rendered batches, without clearing the container, so a large score renders "system by system".
     *  clearFirst=true starts a fresh session (clears prior content, resets the counters). Returns the
     *  source-measure index at which the next batch should continue. fromMeasureIndex is informational --
     *  the drawn frontier is tracked internally. Targets the endless vertical-scroll format and
     *  RenderSingleHorizontalStaffline (laid out once, then drawn to the right batch by batch).
     *  `targetNewSystems`, if set, draws that many whole systems this batch (vertical path only; ignored for
     *  the single horizontal staffline, which is one system). */
    private renderAppend(fromMeasureIndex: number, toMeasureIndex: number, clearFirst: boolean = false,
                         targetNewSystems?: number): number {
        if (!this.graphic) {
            throw new Error("OSMD: load() needs to be called before renderNext()");
        }
        if (this.rules.RenderSingleHorizontalStaffline) {
            // Single horizontal staffline: laid out once like render() (so without the lazy caches and gates of
            // LazyConsistentGraphic), drawing only the newly-entered measures each batch.
            this.rules.LazyConsistentGraphic = false;
            return this.renderAppendGrowingHorizontal(toMeasureIndex, clearFirst);
        }
        // Lazy rendering lays out the whole prefix [first..toMeasureIndex] into one growing, globally-consistent
        // graphic and draws only the newly-stable systems (renderAppendGrowing). LazyConsistentGraphic gates
        // that path's reuse caches; a normal render() forces it off so the caches never affect a non-lazy render.
        this.rules.LazyConsistentGraphic = true;
        return this.renderAppendGrowing(fromMeasureIndex, toMeasureIndex, clearFirst, targetNewSystems);
    }

    /**
     * Incrementally render the loaded sheet one batch at a time, appending each batch so a large score
     * paints progressively ("system by system") instead of blocking on a full render(). The first call --
     * or the first after load(), render() or resetIncrementalRendering() -- starts a fresh session: it
     * clears the container and lays the score out from the first measure to draw. Each later call appends the
     * next batch. Returns progress; once `done` is true all measures are rendered and further calls are no-ops.
     * A session renders the measures render() draws: the whole sheet, or the range set with the options
     * drawFromMeasureNumber and drawUpToMeasureNumber when the session starts (the progress counts the measures in it).
     * Like render(), a session lays the sheet out at the container's width when it starts: to adapt it to a new
     * width, e.g. after a resize, start a new session.
     * Like render(), it applies a changed transposition (osmd.Sheet.Transpose, Instrument.Transpose, osmd.TransposeCalculator)
     * without updateGraphic(): during a session, the next call then starts it over from its first measure.
     * Page labels are drawn with the batch that finalizes their position: the title block (title, subtitle,
     * composer, lyricist) with the first batch, the copyright (below the last system) with the final batch.
     *
     * Pair with {@link enableIncrementalRenderingOnScroll} for scroll-to-load, or {@link renderRemaining}
     * to finish synchronously (e.g. before PDF/image export). Works for the endless vertical-scroll page
     * format and for RenderSingleHorizontalStaffline (one staffline scrolling right).
     *
     * In a container without a width (e.g. a hidden one), no session starts: like render(), the call is skipped with a
     * warning, and returns `done: true` with 0 measures, so that a loop over the batches ends. With autoResize, render()
     * draws the whole sheet once the container has a width.
     *
     * @param options batch options; defaults to 8 visual measures (a multi-rest counts as one). Pass
     *   `systems` instead to advance by whole music systems (vertical only); see {@link IRenderNextOptions}.
     */
    public renderNext(options?: IRenderNextOptions): IRenderNextResult {
        if (this.destroyed) {
            throw this.createDestroyedError("renderNext");
        }
        if (!this.graphic) {
            throw new Error("OSMD: load() needs to be called before renderNext()");
        }
        // A session can't start in a container without a width (see skipsRenderWithoutWidth()). There is no batch to ask
        //   for then (done), so that a loop over the batches ends; with autoResize, render() draws the sheet when the
        //   container gets a width, as it does after a resize during a session.
        if (!this.lazyIncrementalActive && this.skipsRenderWithoutWidth("renderNext")) {
            return { done: true, renderedMeasures: 0, totalMeasures: 0, lastRenderedMeasure: [], nextUnrenderedMeasure: [] };
        }
        this.ensureGraphicPrepared();
        const batchMeasures: number = Math.max(1, options?.measures ?? 8);
        // `systems` (vertical only) advances the frontier by whole music systems instead of measures. A single
        // horizontal staffline is one system, so it ignores `systems` and uses `measures`.
        const systemsOpt: number = options?.systems ?? 0;
        const targetNewSystems: number = systemsOpt > 0 && !this.rules.RenderSingleHorizontalStaffline
            ? Math.max(1, Math.floor(systemsOpt)) : undefined;

        // As in render(): create the graphic again after a change of the transposition or of the visible voices, or for the draw range
        //   of the session (that of render() for a new one) if it would divide the multiple rest measures otherwise.
        const drawRange: [number, number] = this.lazyIncrementalActive
            ? [this.lazyFirstMeasureIndex, this.lazyLastMeasureIndex] : this.getDrawRangeIndices();
        const graphicOutdated: boolean = this.graphicOutdated(drawRange);
        if (graphicOutdated) {
            this.createGraphic(drawRange);
            this.ensureGraphicPrepared();
        }
        const begin: boolean = !this.lazyIncrementalActive;
        if (begin) {
            this.beginIncrementalRendering();
        } else if (graphicOutdated) {
            // start the session over: the systems drawn show the previous transposition or voices
            this.lazyNextSourceIndex = this.lazyFirstMeasureIndex;
        }
        // the session renders the measures [first, last], those of the draw range (see beginIncrementalRendering()), and draws
        //   them up to the last one of the layout (see lazyLayoutLastMeasureIndex)
        const lastMeasureIndex: number = this.lazyLastMeasureIndex;
        // All measures rendered. A draw range without measures has none to render, but the session's first batch lays it out
        //   like render(), which removes the previous drawing from the container.
        if (this.lazyNextSourceIndex > lastMeasureIndex && !begin) {
            const measuresDrawn: number = this.visualMeasureCount(this.lazyFirstMeasureIndex, this.lazyLayoutLastMeasureIndex + 1);
            return {
                done: true, renderedMeasures: measuresDrawn, totalMeasures: measuresDrawn,
                lastRenderedMeasure: this.graphicalMeasuresAtOrBefore(this.lazyLayoutLastMeasureIndex), nextUnrenderedMeasure: []
            };
        }

        const fromMeasureIndex: number = this.lazyNextSourceIndex;
        // In systems mode, seed the layout at the current frontier and let renderAppendGrowing grow it to
        // exactly `targetNewSystems` whole systems; otherwise advance the frontier by the visual-measure count.
        const toMeasureIndex: number = targetNewSystems !== undefined
            ? fromMeasureIndex
            : Math.min(this.visualBatchEndIndex(fromMeasureIndex, batchMeasures), lastMeasureIndex);
        this.lazyNextSourceIndex = this.renderAppend(fromMeasureIndex, toMeasureIndex, begin || graphicOutdated, targetNewSystems);
        this.systemVirtualization.refresh();

        const done: boolean = this.lazyNextSourceIndex > lastMeasureIndex;
        // counted after the batch: its layout can end before the last measure (see lazyLayoutLastMeasureIndex)
        const lastDrawnMeasureIndex: number = Math.min(this.lazyNextSourceIndex - 1, this.lazyLayoutLastMeasureIndex);
        const totalMeasures: number = this.visualMeasureCount(this.lazyFirstMeasureIndex, this.lazyLayoutLastMeasureIndex + 1);
        const renderedMeasures: number = done ? totalMeasures : this.visualMeasureCount(this.lazyFirstMeasureIndex, lastDrawnMeasureIndex + 1);
        this.useWarmKeeper();
        if (this.OnRender) {
            this.OnRender({ incremental: true, done });
        }
        return {
            done, renderedMeasures, totalMeasures,
            lastRenderedMeasure: this.graphicalMeasuresAtOrBefore(lastDrawnMeasureIndex),
            nextUnrenderedMeasure: done ? [] : this.graphicalMeasuresAtOrBefore(this.lazyNextSourceIndex)
        };
    }

    /**
     * Finish an in-progress incremental render synchronously: render all remaining measures at once. Useful
     * before PDF/image export or printing, which need every system, not just the ones scrolled into view.
     * No-op if no incremental render is in progress, or it is already complete.
     */
    public renderRemaining(): void {
        if (!this.lazyIncrementalActive) {
            return;
        }
        const total: number = this.sheet.SourceMeasures.length;
        let guard: number = 0;
        while (!this.IncrementalRenderingComplete && guard++ < total + 2) {
            const before: number = this.lazyNextSourceIndex;
            this.renderNext({ measures: total }); // one big batch -> final batch -> done
            if (this.lazyNextSourceIndex <= before) {
                break; // safety: no forward progress
            }
        }
    }

    /** Whether an incremental render ({@link renderNext}) is in progress (started and not yet reset). */
    public get IncrementalRenderingActive(): boolean {
        return this.lazyIncrementalActive;
    }

    /** Whether the in-progress incremental render has rendered all its measures (up to the last one of the sheet, or of
     *  the draw range, see {@link renderNext}). */
    public get IncrementalRenderingComplete(): boolean {
        return this.lazyIncrementalActive && !!this.sheet && this.lazyNextSourceIndex > this.lazyLastMeasureIndex;
    }

    /** Current incremental-render progress as a snapshot (same shape {@link renderNext} returns), queryable
     *  at any time -- e.g. for a progress bar, or to re-render the same extent after a resize. Reports zero
     *  rendered measures when no session is active (of the measures a session would render now). */
    public get IncrementalRenderProgress(): IRenderNextResult {
        const active: boolean = !!this.graphic && this.lazyIncrementalActive;
        let totalMeasures: number = 0;
        if (this.graphic) {
            // the measures of the draw range, the session's ones if active, up to the last one of its layout (see
            //   beginIncrementalRendering() and lazyLayoutLastMeasureIndex)
            const [firstIndex, lastIndex]: [number, number] = active
                ? [this.lazyFirstMeasureIndex, this.lazyLayoutLastMeasureIndex] : this.getDrawRangeIndices();
            totalMeasures = this.visualMeasureCount(Math.min(firstIndex, this.sheet.SourceMeasures.length - 1), lastIndex + 1);
        }
        const done: boolean = this.IncrementalRenderingComplete;
        const drawnToIndexExcl: number = Math.min(this.lazyNextSourceIndex, this.lazyLayoutLastMeasureIndex + 1);
        const renderedMeasures: number = active ? this.visualMeasureCount(this.lazyFirstMeasureIndex, drawnToIndexExcl) : 0;
        return {
            done, renderedMeasures, totalMeasures,
            lastRenderedMeasure: active && this.lazyNextSourceIndex > this.lazyFirstMeasureIndex
                ? this.graphicalMeasuresAtOrBefore(drawnToIndexExcl - 1) : [],
            nextUnrenderedMeasure: active && !done ? this.graphicalMeasuresAtOrBefore(this.lazyNextSourceIndex) : []
        };
    }

    /** All GraphicalMeasures (one per staff/instrument) at the measure position for source-measure index
     *  `sourceIndex`, walking BACK over collapsed multi-rest members (which have no graphical measure of their
     *  own) to the nearest real measure, not before the first measure of the incremental render. Empty if there is
     *  none at/before it. Used to resolve the last- / next-measure handles for renderNext(). */
    private graphicalMeasuresAtOrBefore(sourceIndex: number): GraphicalMeasure[] {
        if (!this.graphic) {
            return [];
        }
        const measureList: GraphicalMeasure[][] = this.graphic.MeasureList;
        let li: number = Math.min(sourceIndex, measureList.length - 1);
        while (li >= this.lazyFirstMeasureIndex && !(measureList[li] && measureList[li].some(measure => !!measure))) {
            li--;
        }
        return li >= this.lazyFirstMeasureIndex ? measureList[li].filter(measure => !!measure) : [];
    }

    /**
     * Abandon any in-progress incremental render, so that the next renderNext() starts a new session (with the draw
     * range set then). Called automatically by render(), load() and clear(); call it directly only to cancel a session
     * yourself.
     */
    public resetIncrementalRendering(): void {
        if (!this.lazyIncrementalActive) {
            return;
        }
        this.disableIncrementalRenderingOnScroll(); // drop the scroll listener with the session
        this.lazyIncrementalActive = false;
        this.lazyNextSourceIndex = 0;
        this.lazyDrawnSystemCount = 0;
        this.lazyDrawnSystemPositions = [];
        this.lazyDrawnHMeasureCount = 0;
    }

    /**
     * Drive {@link renderNext} automatically from scrolling: render the next batch whenever the user scrolls
     * within ~1.5 viewports of the not-yet-rendered edge (the page bottom for the endless vertical format,
     * the right edge for RenderSingleHorizontalStaffline). Renders the first batch immediately if none has
     * been rendered yet, then keeps appending as the user scrolls, and detaches itself once all measures
     * are rendered (see {@link renderNext}). Re-enabling replaces any previous listener; reset/render/load also detach it.
     *
     * @param options batch size (as {@link renderNext}) plus an optional `scrollElement` -- the element whose
     *  scrolling drives loading. Defaults to the OSMD container for a single horizontal staffline (it scrolls
     *  horizontally) and to `window` otherwise (the page scrolls vertically).
     */
    public enableIncrementalRenderingOnScroll(options?: IRenderNextOptions & { scrollElement?: HTMLElement | Window }): void {
        if (typeof window === "undefined") {
            return; // no DOM (e.g. headless): nothing to attach to
        }
        this.disableIncrementalRenderingOnScroll(); // drop any previous listener
        const horizontal: boolean = this.rules.RenderSingleHorizontalStaffline;
        const target: HTMLElement | Window = options?.scrollElement ?? (horizontal ? this.container : window);
        const batchOptions: IRenderNextOptions = { measures: options?.measures };
        if (!this.IncrementalRenderingActive) {
            this.renderNext(batchOptions); // paint the first batch so there is something to scroll
            if (!this.IncrementalRenderingActive) {
                return; // skipped: the container has no width (see renderNext())
            }
        }
        if (this.IncrementalRenderingComplete) {
            return; // all measures fit in the first batch; nothing to load on scroll
        }
        const nearEnd: () => boolean = () => {
            const margin: number = 1.5; // start loading ~1.5 viewports before the edge
            const el: HTMLElement = (target === window ? document.scrollingElement || document.documentElement : target) as HTMLElement;
            if (horizontal) {
                return el.scrollLeft + el.clientWidth >= el.scrollWidth - el.clientWidth * margin;
            }
            if (target === window) {
                return window.scrollY + window.innerHeight >= document.body.scrollHeight - window.innerHeight * margin;
            }
            return el.scrollTop + el.clientHeight >= el.scrollHeight - el.clientHeight * margin;
        };
        let loading: boolean = false;
        const onScroll: () => void = () => {
            if (loading || !this.IncrementalRenderingActive || this.IncrementalRenderingComplete) {
                return;
            }
            if (!nearEnd()) {
                return;
            }
            loading = true;
            // Appending the next batch below the viewport must not move it. Each batch resizes the backend,
            // and on a reconciliation batch createOrRefreshRenderBackend() momentarily removes the rendered
            // SVG/Canvas from the container (collapsing the page height) before the new, taller content is
            // sized in. While the height is collapsed, some mobile/touch browsers clamp the scroll offset to
            // the top; desktop browsers keep it. So capture the offset before the batch and restore it after
            // -- the in-core equivalent of renderAndScrollBack() for an incremental append. Unlike
            // renderAndScrollBack() (which restores by PERCENT, suited to a same-height re-render), here the
            // page grows, so we restore the ABSOLUTE offset: the appended systems are below the viewport and
            // must not shift it. The !== guards make it a no-op when nothing moved (e.g. on desktop).
            const scrollEl: HTMLElement = (target === window
                ? (document.scrollingElement || document.documentElement)
                : target) as HTMLElement;
            const savedTop: number = scrollEl.scrollTop;
            const savedLeft: number = scrollEl.scrollLeft;
            const result: IRenderNextResult = this.renderNext(batchOptions);
            if (scrollEl.scrollTop !== savedTop) {
                scrollEl.scrollTop = savedTop;
            }
            if (scrollEl.scrollLeft !== savedLeft) {
                scrollEl.scrollLeft = savedLeft;
            }
            loading = false;
            if (result.done) {
                this.disableIncrementalRenderingOnScroll();
            } else if (typeof window.requestAnimationFrame === "function") {
                window.requestAnimationFrame(onScroll); // keep filling while still near the edge
            }
        };
        this.lazyScrollTarget = target;
        this.lazyScrollHandler = onScroll;
        target.addEventListener("scroll", onScroll, { passive: true });
        if (typeof window.requestAnimationFrame === "function") {
            window.requestAnimationFrame(onScroll); // fill the viewport if the first batch was short
        }
    }

    /** Stop driving {@link renderNext} from scrolling (see {@link enableIncrementalRenderingOnScroll}). */
    public disableIncrementalRenderingOnScroll(): void {
        if (this.lazyScrollHandler && this.lazyScrollTarget) {
            this.lazyScrollTarget.removeEventListener("scroll", this.lazyScrollHandler);
        }
        this.lazyScrollHandler = undefined;
        this.lazyScrollTarget = undefined;
    }

    /**
     * Keep only systems near the scroll viewport attached to the live SVG DOM. SVG backend only.
     * Detached systems retain their exact nodes, including colors, opacity, and event listeners.
     */
    public enableSystemVirtualization(options?: ISystemVirtualizationOptions): void {
        this.systemVirtualization.enable(options);
    }

    /** Stop viewport virtualization and, by default, put every rendered system back in the SVG. */
    public disableSystemVirtualization(restore: boolean = true): void {
        this.systemVirtualization.disable(restore);
    }

    /** Be notified when virtualization draws, reattaches or detaches systems, instead of observing the DOM. */
    public addSystemLifecycleListener(listener: SystemLifecycleListener): () => void {
        return this.systemVirtualization.addLifecycleListener(listener);
    }

    /** Re-evaluate the active system window after an application-controlled scroll or layout change. */
    public updateSystemVirtualization(): void {
        this.systemVirtualization.refresh();
    }

    public get SystemVirtualizationStats(): ISystemVirtualizationStats {
        return this.systemVirtualization.stats;
    }

    /** Start a fresh incremental session for the measures render() draws (the draw range). Lazy needs the greedy,
     *  forward-stable builder so earlier systems keep their position as the prefix grows -- public OSMD is
     *  greedy-only, so that is already in effect (nothing to force). */
    private beginIncrementalRendering(): void {
        this.lazyIncrementalActive = true;
        const [firstIndex, lastIndex]: [number, number] = this.getDrawRangeIndices();
        const lastSheetMeasureIndex: number = this.sheet.SourceMeasures.length - 1;
        // a first measure beyond the sheet is its last one, as in the layout (see MusicSheetCalculator.calculateMusicSystems())
        this.lazyFirstMeasureIndex = Math.min(firstIndex, lastSheetMeasureIndex);
        this.lazyLastMeasureIndex = Math.min(lastIndex, lastSheetMeasureIndex);
        this.lazyLayoutLastMeasureIndex = this.lazyLastMeasureIndex;
        this.lazyNextSourceIndex = this.lazyFirstMeasureIndex;
        // The draw range is in the rules like after render(), e.g. for the cursor. The batches lay out the measures from
        //   its first one up to theirs, and leave the rules as they were.
        this.rules.MinMeasureToDrawIndex = this.lazyFirstMeasureIndex;
        this.rules.MaxMeasureToDrawIndex = lastIndex;
    }

    /** Number of VISUAL measures in the source-measure range [from, toExcl): a multi-rest renders as one
     *  GraphicalMeasure when RenderMultipleRestMeasures collapses it, so it counts as one. */
    private visualMeasureCount(from: number, toExcl: number): number {
        const sourceMeasures: SourceMeasure[] = this.sheet.SourceMeasures;
        const collapse: boolean = this.rules.RenderMultipleRestMeasures;
        let visual: number = 0;
        let i: number = from;
        while (i < toExcl && i < sourceMeasures.length) {
            i += collapse && sourceMeasures[i].multipleRestMeasures > 0 ? sourceMeasures[i].multipleRestMeasures : 1;
            visual++;
        }
        return visual;
    }

    /** Source-measure index (inclusive) at which a batch of `visualCount` visual measures starting at source
     *  index `from` ends -- i.e. the toMeasureIndex to render this batch (multi-rests collapsed; see above). */
    private visualBatchEndIndex(from: number, visualCount: number): number {
        const sourceMeasures: SourceMeasure[] = this.sheet.SourceMeasures;
        const collapse: boolean = this.rules.RenderMultipleRestMeasures;
        let end: number = from;
        let visual: number = 0;
        while (visual < visualCount && end < sourceMeasures.length) {
            end += collapse && sourceMeasures[end].multipleRestMeasures > 0 ? sourceMeasures[end].multipleRestMeasures : 1;
            visual++;
        }
        return Math.min(end - 1, sourceMeasures.length - 1);
    }

    /**
     * Lazy rendering (EngravingRules.LazyConsistentGraphic). Lays out the WHOLE prefix
     * [first..toMeasureIndex] of the draw range (see beginIncrementalRendering()) into ONE globally-consistent
     * graphic and draws only the systems that became stable since the previous batch, appending them into the
     * shared backend below what's already drawn.
     *
     * There is no per-batch vertical offset or seam-distance
     * computation: the real Y-layout already stacks and crisp-snaps every system at its final absolute
     * position, and the title space / first-system instrument names fall out of laying the prefix from
     * its first measure each time. The greedy builder is forward-stable -- every system except the LAST of a
     * prefix is *usually* forward-stable (an interior system keeps its position as the prefix grows), so
     * we skip drawing the already-drawn systems above and DEFER the last system of a non-final batch (it
     * is unstretched and shifts once it becomes an interior, stretched system next batch). But this is NOT
     * universal -- some scores re-position earlier systems by several px as later systems are added -- so
     * each batch first VERIFIES the drawn systems against their drawn Y (lazyDrawnSystemPositions) and, if any
     * moved, redraws the whole drawn range at the corrected positions (reconciliation). In the common
     * (stable) case nothing is redrawn. See export/inspect_prefix_stability.mjs / inspect_optionb_drawpos.mjs.
     * The layout can have several pages (NewPageAtXMLNewPageAttribute), each drawn into its own backend like in render():
     * the systems are counted through the pages, and a backend is added for each page the layout grows to.
     *
     * @returns the source-measure index at which the next batch should continue (past the last drawn system).
     */
    private renderAppendGrowing(fromMeasureIndex: number, toMeasureIndex: number, clearFirst: boolean,
                                targetNewSystems?: number): number {
        if (clearFirst) {
            this.lazyDrawnSystemCount = 0;
            this.lazyDrawnSystemPositions = [];
            this.graphic.GetCalculator?.clearSkyBottomLineCache(); // fresh lazy session: drop reused sky/bottom lines
            this.lazyLayoutWidth = this.getContainerContentWidth();
            this.layoutContainerWidth = this.lazyLayoutWidth;
            this.lazyLayoutLastMeasureIndex = this.lazyLastMeasureIndex; // until a layout ends before it
        }
        const lastMeasureIndex: number = this.lazyLastMeasureIndex; // of the draw range

        this.sheet.pageWidth = this.lazyLayoutWidth / this.zoom / 10.0; // the session's width, see lazyLayoutWidth
        this.rules.PageHeight = 100001; // lazy assumes the endless (vertical scroll) page format

        // Lay out the whole prefix [first..to] of the draw range. Earlier systems are re-laid-out identically;
        // only the newly-completed systems are drawn below. The rules' draw range is the prefix only for its layout,
        // then the session's again (see beginIncrementalRendering()).
        // Ensure the prefix yields enough NEW systems to draw beyond the deferred last one (>= drawn +
        // minNewSystems + 1 systems total); otherwise this batch only grew the current last system. Extend the
        // prefix by a batch span and retry until enough systems appear or the draw range ends. minNewSystems is 1 in
        // measures mode (the batch's other complete systems are drawn too) and `targetNewSystems` in systems mode.
        // Then also until the octave shifts reaching into the systems to draw end in a complete system (see
        // lazySpannerEndInLastSystem()), without drawing the systems this adds.
        // A layout ending before the prefix's end at the last system or page to draw (drawUpToSystemNumber, drawUpToPageNumber)
        //   has all systems to draw, and a longer prefix wouldn't add any: the batch lays out the draw range up to its last measure
        //   like render() instead, and is the final one.
        const drawRangeIndices: [number, number] = [this.rules.MinMeasureToDrawIndex, this.rules.MaxMeasureToDrawIndex];
        let systems: MusicSystem[];
        const minNewSystems: number = targetNewSystems ?? 1;
        const extendStep: number = Math.max(4, toMeasureIndex - fromMeasureIndex + 1);
        // Hold the last (unstretched, not-yet-stable) system unless this is the final batch, where the last
        // system is the sheet's true last system and never changes again. In systems mode, also cap the draw at
        // `targetNewSystems` new systems (the layout may hold a few more from the last extend step), so each
        // batch advances by exactly that many whole systems; any extra are re-laid-out and drawn next batch.
        const nonFinalDrawToIdxExcl: (layoutSystemCount: number) => number = (layoutSystemCount: number): number =>
            targetNewSystems !== undefined ? Math.min(layoutSystemCount - 1, this.lazyDrawnSystemCount + targetNewSystems) : layoutSystemCount - 1;
        let drawToIdxExclCap: number; // the draw range end once the layout has enough systems
        let extendedTo: number = Math.min(toMeasureIndex, lastMeasureIndex);
        try {
            this.rules.MinMeasureToDrawIndex = this.lazyFirstMeasureIndex;
            this.rules.MaxMeasureToDrawIndex = extendedTo;
            this.graphic.reCalculate();
            systems = this.lazyLaidOutSystems();
            while (systems.length > 0 && extendedTo < lastMeasureIndex) {
                let extendTo: number = extendedTo + extendStep;
                if (systems.length >= this.lazyDrawnSystemCount + minNewSystems + 1) {
                    if (drawToIdxExclCap === undefined) {
                        drawToIdxExclCap = nonFinalDrawToIdxExcl(systems.length);
                    }
                    const spannerEndIndex: number = this.lazySpannerEndInLastSystem(systems, drawToIdxExclCap);
                    if (spannerEndIndex < 0) {
                        break; // long enough
                    }
                    extendTo = Math.max(extendTo, spannerEndIndex + 1);
                }
                if (this.lazyLayoutEndBefore(systems, extendedTo) !== undefined) {
                    extendTo = lastMeasureIndex; // all systems to draw are laid out
                }
                extendedTo = Math.min(extendTo, lastMeasureIndex);
                this.rules.MaxMeasureToDrawIndex = extendedTo;
                this.graphic.reCalculate();
                systems = this.lazyLaidOutSystems();
            }
        } finally {
            [this.rules.MinMeasureToDrawIndex, this.rules.MaxMeasureToDrawIndex] = drawRangeIndices;
        }
        if (systems.length === 0) {
            // Nothing laid out, e.g. with no instrument visible or a draw range without measures: no page. A batch starting the
            //   session removes the previous drawing from the container like render(), which creates a backend for each page.
            if (clearFirst) {
                this.createOrRefreshRenderBackend(this.lazyLayoutWidth);
            }
            return lastMeasureIndex + 1; // nothing more to do
        }
        const layoutEndIndex: number | undefined = this.lazyLayoutEndBefore(systems, extendedTo);
        if (layoutEndIndex !== undefined) {
            this.lazyLayoutLastMeasureIndex = layoutEndIndex; // the session's progress counts up to it
        }

        const systemCount: number = systems.length;
        const finalBatch: boolean = extendedTo >= lastMeasureIndex;
        const sysAbsY: (i: number) => number = i => systems[i].StaffLines[0].PositionAndShape.AbsolutePosition.y;
        // Reconciliation: forward-stability is not universal -- this batch's full-prefix layout may have
        // moved an already-drawn system. If the topmost drawn system whose Y (or page) changed exists, everything
        // already drawn is stale; clear the backends and redraw the whole drawn range at the corrected
        // positions. In the common (stable) case nothing moved and we only append the new systems.
        let someDrawnSystemMoved: boolean = false;
        for (let i: number = 0; i < this.lazyDrawnSystemCount && i < systemCount; i++) {
            const drawnAt: { pageNumber: number, y: number } = this.lazyDrawnSystemPositions[i];
            if (drawnAt && (drawnAt.pageNumber !== systems[i].Parent.PageNumber || Math.abs(sysAbsY(i) - drawnAt.y) > 1e-4)) {
                someDrawnSystemMoved = true;
                break;
            }
        }
        // Recreating the backend erases everything, so a recreate batch redraws from system 0.
        const recreateBackend: boolean = clearFirst || someDrawnSystemMoved;
        const drawFromIdx: number = recreateBackend ? 0 : this.lazyDrawnSystemCount;
        const drawToIdxExcl: number = finalBatch ? systemCount : Math.min(nonFinalDrawToIdxExcl(systemCount), drawToIdxExclCap ?? systemCount);
        if (drawToIdxExcl <= drawFromIdx) {
            this.lazyDrawnSystemCount = Math.max(this.lazyDrawnSystemCount, drawToIdxExcl);
            return lastMeasureIndex + 1; // no new complete system (only happens at the very end)
        }

        // createOrRefreshRenderBackend rebuilds the drawer (resetting the lazy draw window), so it must
        // run BEFORE setting that window below. A purely-appending batch keeps the existing backends.
        // The backends are as wide as the layout, also when the container's width has changed since the first batch.
        if (recreateBackend) {
            this.createOrRefreshRenderBackend(this.lazyLayoutWidth);
        }
        // A backend per page, like render(): a batch whose layout reaches a new page adds one for it. Grow each backend to fit
        // through the last system we draw on its page. Once all of a page's systems are drawn (on the final batch, or before
        // a page break), size it exactly as createOrRefreshRenderBackend() does for a full page so the finished image height
        // byte-matches a normal render; until then, size it to its last drawn system (the deferred system's slot is filled
        // next batch), or to nothing, as long as none of its systems is drawn.
        const pageWidth: number = this.drawer.Backends[0].width;
        let pageFirstSystemIdx: number = 0; // index of the page's first system in systems
        for (const page of this.graphic.MusicPages) {
            if (page.PageNumber > this.rules.MaxPageToDrawNumber) {
                break; // not drawn, see lazyLaidOutSystems()
            }
            let backend: VexFlowBackend = this.drawer.Backends[page.PageNumber - 1];
            if (!backend) {
                backend = this.createBackend(this.backendType, page);
                this.drawer.Backends.push(backend);
            }
            const isCanvas: boolean = backend.getOSMDBackendType() === BackendType.Canvas;
            const drawnPageSystems: number = Math.min(Math.max(drawToIdxExcl - pageFirstSystemIdx, 0), page.MusicSystems.length);
            let heightUnits: number = 0;
            if (drawnPageSystems === page.MusicSystems.length) {
                heightUnits = page.PositionAndShape.Size.height + this.rules.PageBottomMargin + page.PositionAndShape.BorderTop;
            } else if (drawnPageSystems > 0) {
                const lastDrawnSystem: MusicSystem = page.MusicSystems[drawnPageSystems - 1];
                heightUnits = lastDrawnSystem.PositionAndShape.AbsolutePosition.y
                    + lastDrawnSystem.PositionAndShape.BorderBottom + this.rules.PageBottomMargin;
            }
            if (heightUnits > 0) {
                if (isCanvas) {
                    heightUnits += 0.1; // Canvas bug: cuts off the bottom pixel with PageBottomMargin = 0
                }
                if (this.rules.RenderTitle) {
                    heightUnits += this.rules.TitleTopDistance; // title sits above the first system
                }
            }
            backend.graphicalMusicPage = page;
            backend.resize(pageWidth, heightUnits * 10 * this.zoom);
            if (pageFirstSystemIdx >= drawFromIdx) {
                // Nothing drawn on the page yet (e.g. a new backend): its background, as createOrRefreshRenderBackend() sets it.
                //   (A canvas keeps it when resized only from a size greater than 0.)
                backend.clear();
            }
            // Re-establish the default music color: createOrRefreshRenderBackend sets it on the first batch,
            // but a reused canvas backend keeps stateful context and could inherit a stale fill/stroke color.
            backend.getContext().setFillStyle(this.rules.DefaultColorMusic);
            backend.getContext().setStrokeStyle(this.rules.DefaultColorMusic);
            pageFirstSystemIdx += page.MusicSystems.length;
        }
        this.drawer.setZoom(this.zoom);

        if (this.drawingParameters.drawCursors) {
            this.graphic.Cursors.length = 0; // clear any stale graphical cursors before drawing, as render() does
        }
        // Mark everything up to the drawn range as on-screen (for playback/cursor lookups), since the
        // graphic was rebuilt this batch.
        for (let i: number = 0; i < drawToIdxExcl; i++) {
            for (const staffLine of systems[i].StaffLines) {
                for (const measure of staffLine.Measures) {
                    if (measure?.parentSourceMeasure) { // some graphical measures (e.g. extra-instruction) have none
                        measure.parentSourceMeasure.WasRendered = true;
                    }
                }
            }
        }
        // Draw only the new systems (and the title block only on the first batch), counted through the pages; see drawPage().
        this.drawer.LazyDrawSystemsFromIndex = drawFromIdx;
        this.drawer.LazyDrawSystemsToIndexExcl = drawToIdxExcl;
        this.drawer.drawSheet(this.graphic);
        this.drawer.LazyDrawSystemsFromIndex = -1;
        this.drawer.LazyDrawSystemsToIndexExcl = Number.POSITIVE_INFINITY;

        // Reposition the HTML cursors for this batch's growing / reconciled layout, mirroring render()'s
        // post-draw cursor handling. When the backend was rebuilt this batch, re-create the cursors on it
        // (enableOrDisableCursors restores their position via RestoreCursorAfterRerender); otherwise just
        // update() them in place. update() no-ops if the cursor's target measure isn't laid out yet (the
        // user hasn't scrolled there) -- a later batch repositions it once that measure is rendered.
        if (this.drawingParameters.drawCursors) {
            if (recreateBackend) {
                this.enableOrDisableCursors(this.drawingParameters.drawCursors);
            }
            this.cursors.forEach(cursor => cursor.update());
        }

        // Record where each drawn system landed, so the next batch can detect (and reconcile) any that
        // the growing layout moves.
        for (let i: number = 0; i < drawToIdxExcl; i++) {
            this.lazyDrawnSystemPositions[i] = { pageNumber: systems[i].Parent.PageNumber, y: sysAbsY(i) };
        }
        this.lazyDrawnSystemPositions.length = drawToIdxExcl;
        this.lazyDrawnSystemCount = drawToIdxExcl;
        this.rules.RenderCount++;

        if (finalBatch) {
            return lastMeasureIndex + 1;
        }
        // Continue at the deferred (held) system's first source measure.
        const heldMeasures: GraphicalMeasure[] = systems[drawToIdxExcl].StaffLines[0].Measures;
        return this.sheet.SourceMeasures.indexOf(heldMeasures[0].parentSourceMeasure);
    }

    /**
     * Lazy rendering: where the furthest octave shift ends that reaches into the systems a batch draws and ends in the last
     * system of its layout, or beyond. Such an octave shift needs a longer layout: its parts depend on the system it ends in,
     * e.g. it ends at a note there. The last system still grows (and is stretched once complete), and the end can be beyond the
     * layout, which cuts it off there, so the systems drawn now would be drawn with other parts. They aren't drawn again. An
     * octave shift without an end is ignored. A wedge doesn't need a longer layout: its part in a system depends only on that
     * system and the wedge's first one (see MusicSheetCalculator.calculateGraphicalContinuousDynamic()).
     * @param systems the systems of the layout, in order
     * @param drawToIdxExcl the index of the first system the batch doesn't draw
     * @returns the source-measure index of the furthest end in the last system or beyond, or -1 if there is none
     */
    private lazySpannerEndInLastSystem(systems: MusicSystem[], drawToIdxExcl: number): number {
        const firstMeasureIndex: (system: MusicSystem) => number =
            (system: MusicSystem): number => system.StaffLines[0].Measures[0].parentSourceMeasure.measureListIndex;
        const firstUndrawnMeasureIndex: number = firstMeasureIndex(systems[drawToIdxExcl]);
        let endIndex: number = -1;
        for (let measureIndex: number = 0; measureIndex < firstUndrawnMeasureIndex; measureIndex++) {
            for (const staffExpressions of this.sheet.SourceMeasures[measureIndex].StaffLinkedExpressions) {
                for (const multiExpression of staffExpressions) {
                    const end: MultiExpression = multiExpression.OctaveShiftStart?.ParentEndMultiExpression;
                    if (end) {
                        endIndex = Math.max(endIndex, end.SourceMeasureParent.measureListIndex);
                    }
                }
            }
        }
        return endIndex >= firstMeasureIndex(systems[systems.length - 1]) ? endIndex : -1;
    }

    /**
     * Lazy rendering: the systems of the laid-out pages that are drawn (see EngravingRules.MaxPageToDrawNumber), in order.
     * The systems a batch draws are counted through them (see MusicSheetDrawer.LazyDrawSystemsFromIndex).
     * @returns the systems of the pages in order
     */
    private lazyLaidOutSystems(): MusicSystem[] {
        const systems: MusicSystem[] = [];
        for (const page of this.graphic.MusicPages) {
            if (page.PageNumber > this.rules.MaxPageToDrawNumber) {
                break; // not drawn (see VexFlowMusicSheetDrawer.drawSheet())
            }
            systems.push(...page.MusicSystems);
        }
        return systems;
    }

    /**
     * Lazy rendering: where the layout ends if it ends before a measure at the last system or page to draw
     * (EngravingRules.MaxSystemToDrawNumber, MaxPageToDrawNumber: drawUpToSystemNumber, drawUpToPageNumber). Then the layout has
     * all systems to draw, and the measures after them aren't drawn, like in render().
     * @param systems the systems of the layout that are drawn, in order (see lazyLaidOutSystems())
     * @param measureIndex the source-measure index of the last measure laid out
     * @returns the source-measure index of the last measure of the last system (of a multiple rest measure there: its last
     *  one), or undefined if the layout doesn't end before measureIndex at the last system or page to draw
     */
    private lazyLayoutEndBefore(systems: MusicSystem[], measureIndex: number): number | undefined {
        const pages: GraphicalMusicPage[] = this.graphic.MusicPages;
        let systemCount: number = 0;
        for (const page of pages) {
            systemCount += page.MusicSystems.length;
        }
        const lastPageNumber: number = pages.length > 0 ? pages[pages.length - 1].PageNumber : 0;
        if (systemCount < this.rules.MaxSystemToDrawNumber && lastPageNumber <= this.rules.MaxPageToDrawNumber) {
            return undefined; // all systems and pages laid out are drawn
        }
        const measures: GraphicalMeasure[] = systems[systems.length - 1]?.StaffLines[0]?.Measures ?? [];
        // the last one with a source measure (an extra instruction measure ending a system has none)
        const lastSourceMeasure: SourceMeasure = measures.map(measure => measure.parentSourceMeasure).filter(sourceMeasure => sourceMeasure).pop();
        if (!lastSourceMeasure) {
            return undefined;
        }
        const restMeasures: number = this.rules.RenderMultipleRestMeasures && lastSourceMeasure.multipleRestMeasures > 0
            ? lastSourceMeasure.multipleRestMeasures : 1; // see visualMeasureCount()
        const endIndex: number = lastSourceMeasure.measureListIndex + restMeasures - 1;
        return endIndex < measureIndex ? endIndex : undefined;
    }

    /**
     * Lazy rendering for RenderSingleHorizontalStaffline (one staffline growing to the right, horizontal scroll).
     * The first batch lays the score (its draw range) out like render(), and each batch reuses that layout and draws only the
     * measures entering the drawn frontier (and the elements ending in them), so that the finished incremental render
     * looks like a normal render(). Laying out a growing prefix instead would move what is drawn already, e.g. a
     * higher note further right moves the staffline down, and the lyrics of a staffline are aligned at one height.
     * Laying out the whole staffline is fast anyway, drawing it is the expensive part, which stays lazy.
     * The layout can have several systems, e.g. with forced system breaks (NewSystemAtXMLNewSystemAttribute,
     * RenderXMeasuresPerLineAkaSystem) or where the staffline would be wider than SheetMaximumWidth, and pages
     * (NewPageAtXMLNewPageAttribute): the frontier advances through the measures of all systems in order, and each
     * system a batch reaches is drawn in its own x-window, from the system's previous frontier to its new one. A
     * system's first window is open to the left and its last one to the right, so that every element is drawn once,
     * also the ones beyond the first or last measure, like the instrument names or a long last chord symbol.
     * @param toMeasureIndex source-measure index up to which this batch draws (at least the next measure, so that
     *  the frontier always advances)
     * @param clearFirst whether this batch starts a new session: lays the score out and creates the backends
     * @returns the source-measure index at which the next batch should continue.
     */
    private renderAppendGrowingHorizontal(toMeasureIndex: number, clearFirst: boolean): number {
        const lastMeasureIndex: number = this.lazyLastMeasureIndex; // of the draw range
        if (clearFirst) {
            this.lazyDrawnHMeasureCount = 0;
            // The draw range like render() (in the rules, see beginIncrementalRendering()), on one horizontal staffline:
            //   SheetMaximumWidth keeps it a single system (mirrors render()).
            this.sheet.pageWidth = this.rules.SheetMaximumWidth / this.zoom / 10.0;
            this.rules.PageHeight = 100001;
            this.graphic.reCalculate();
            this.createOrRefreshRenderBackend(); // a backend per page, like render()
            // the session's progress counts up to the last measure laid out (see lazyLayoutLastMeasureIndex)
            this.lazyLayoutLastMeasureIndex = this.lazyLayoutEndBefore(this.lazyLaidOutSystems(), lastMeasureIndex) ?? lastMeasureIndex;
        }
        // The measures of the top staffline of all systems, in drawing order.
        const pages: GraphicalMusicPage[] = this.graphic.MusicPages.filter(page => page.PageNumber <= this.rules.MaxPageToDrawNumber);
        const measures: GraphicalMeasure[] = [];
        for (const page of pages) {
            for (const system of page.MusicSystems) {
                measures.push(...(system.StaffLines[0]?.Measures ?? []));
            }
        }
        if (measures.length === 0) {
            return lastMeasureIndex + 1; // nothing laid out, e.g. with no instrument visible: the first batch removed the drawing above
        }
        // Draw the measures up to source measure toMeasureIndex, but at least one. A measure without a source measure
        // (an extra instruction measure ending a system, e.g. with the next system's clef) is drawn with the one before.
        const drawFromIdx: number = Math.min(this.lazyDrawnHMeasureCount, measures.length);
        let drawToIdxExcl: number = drawFromIdx;
        while (drawToIdxExcl < measures.length) {
            const sourceMeasure: SourceMeasure = measures[drawToIdxExcl].parentSourceMeasure;
            if (sourceMeasure && sourceMeasure.measureListIndex > toMeasureIndex && drawToIdxExcl > drawFromIdx) {
                break;
            }
            drawToIdxExcl++;
        }
        const finalBatch: boolean = drawToIdxExcl >= measures.length;

        // The x-window of each system the batch reaches: the objects whose right edge is in (fromX, toX] -- the new
        // measures, and the elements ending in them, e.g. a slur. The windows end a bit right of a measure's right edge.
        const measureRightX: (m: GraphicalMeasure) => number =
            m => m.PositionAndShape.AbsolutePosition.x + m.PositionAndShape.BorderRight;
        const windows: Map<MusicSystem, { fromX: number, toX: number }> = new Map();
        const drawnWidthUnits: Map<GraphicalMusicPage, number> = new Map(); // up to the right edge of the drawn measures
        let systemStartIdx: number = 0; // index of the system's first measure in measures
        for (const page of pages) {
            for (const system of page.MusicSystems) {
                const systemMeasures: GraphicalMeasure[] = system.StaffLines[0]?.Measures ?? [];
                const fromIdx: number = Math.max(drawFromIdx - systemStartIdx, 0);
                const toIdxExcl: number = Math.min(drawToIdxExcl - systemStartIdx, systemMeasures.length);
                systemStartIdx += systemMeasures.length;
                if (toIdxExcl > 0) {
                    const widthUnits: number = measureRightX(systemMeasures[toIdxExcl - 1]) + this.rules.PageRightMargin;
                    drawnWidthUnits.set(page, Math.max(drawnWidthUnits.get(page) ?? 0, widthUnits));
                }
                if (fromIdx < toIdxExcl) {
                    windows.set(system, {
                        fromX: fromIdx === 0 ? Number.NEGATIVE_INFINITY : measureRightX(systemMeasures[fromIdx - 1]) + 1e-4,
                        toX: toIdxExcl === systemMeasures.length ? Number.POSITIVE_INFINITY : measureRightX(systemMeasures[toIdxExcl - 1]) + 1e-4
                    });
                }
            }
        }

        // Size the backends like render() (createOrRefreshRenderBackend), but until the final batch only as wide as the
        // measures drawn on them, so that the container's scroll width grows with the drawn frontier (scrolling near it
        // renders the next batch, see enableIncrementalRenderingOnScroll()).
        const fullWidthUnits: number = this.rules.PageLeftMargin + this.graphic.MusicPages[0].PositionAndShape.Size.width + this.rules.PageRightMargin;
        for (const backend of this.drawer.Backends) {
            const page: GraphicalMusicPage = backend.graphicalMusicPage;
            const widthUnits: number = finalBatch ? fullWidthUnits : drawnWidthUnits.get(page) ?? 0;
            let heightUnits: number = page.PositionAndShape.Size.height + this.rules.PageBottomMargin + page.PositionAndShape.BorderTop;
            if (backend.getOSMDBackendType() === BackendType.Canvas) {
                heightUnits += 0.1; // Canvas bug: cuts off the bottom pixel with PageBottomMargin = 0
            }
            if (this.rules.RenderTitle) {
                heightUnits += this.rules.TitleTopDistance;
            }
            backend.resize(widthUnits * 10 * this.zoom, heightUnits * 10 * this.zoom);
            backend.getContext().setFillStyle(this.rules.DefaultColorMusic);
            backend.getContext().setStrokeStyle(this.rules.DefaultColorMusic);
        }
        this.drawer.setZoom(this.zoom);

        if (this.drawingParameters.drawCursors) {
            this.graphic.Cursors.length = 0;
        }
        // The page labels (title/credits) and bounding boxes are drawn with the final batch, when the page is drawn to its
        // full width; drawPage() then opens the x-window so none are dropped (see drawPage).
        this.drawer.LazyDrawSystemWindows = windows;
        this.drawer.LazySkipPageLabels = !finalBatch;
        this.drawer.drawSheet(this.graphic);
        this.drawer.LazyDrawSystemWindows = undefined;
        this.drawer.LazySkipPageLabels = false;

        if (this.drawingParameters.drawCursors) {
            if (clearFirst) {
                this.enableOrDisableCursors(this.drawingParameters.drawCursors);
            }
            this.cursors.forEach(cursor => cursor.update());
        }
        this.lazyDrawnHMeasureCount = drawToIdxExcl;
        this.rules.RenderCount++;
        if (finalBatch) {
            return lastMeasureIndex + 1;
        }
        // Continue at the first measure not drawn yet.
        return measures[drawToIdxExcl].parentSourceMeasure.measureListIndex;
    }

    /**
     * Removes this instance's pages from the container: the element of each backend (SVG or canvas), with the cursors in it.
     * Not the container's other children, e.g. an overlay or a loading indicator of the app. The backends are known by
     * reference, so none is missed (before, all children of the container were removed instead, see #678).
     */
    private removeOwnPages(): void {
        if (!this.drawer?.Backends) {
            return;
        }
        for (const backend of this.drawer.Backends) {
            backend.removeFromContainer(this.container);
            backend.free();
        }
        this.drawer.Backends.length = 0;
    }

    /**
     * Removes the pages that other OpenSheetMusicDisplay instances left in the container: its children with the id
     * "osmdCanvasPage" + page number, which every instance gives its page elements. E.g. the drawing of an instance that
     * was disposed (dispose() keeps the drawing) before this one was made for the same container, as a component does
     * that is mounted, unmounted and mounted again (React's StrictMode): the container would show two scores.
     * Called at this instance's first render, after removeOwnPages(); later renders only remove its own pages.
     */
    private removeLeftoverPages(): void {
        for (const child of Array.from(this.container.children)) {
            if (/^osmdCanvasPage\d+$/.test(child.id)) {
                child.remove();
            }
        }
    }

    /**
     * Removes the backends (SVG or canvas) from the container and creates a new one for each page drawn.
     * @param pageWidth the width of the pages in pixels. By default the container's content width, read after removing the
     *  backends. An incremental render passes the width all its batches lay the sheet out at (see lazyLayoutWidth).
     */
    protected createOrRefreshRenderBackend(pageWidth?: number): void {
        // console.log("[OSMD] createOrRefreshRenderBackend()");

        this.systemVirtualization.invalidate();

        // Remove old backends
        this.removeOwnPages();
        if (!this.leftoverPagesRemoved) {
            this.removeLeftoverPages();
            this.leftoverPagesRemoved = true;
        }

        // Create the drawer
        this.drawingParameters.Rules = this.rules;
        this.drawer = new VexFlowMusicSheetDrawer(this.drawingParameters); // note that here the drawer.drawableBoundingBoxElement is lost. now saved in OSMD.
        this.drawer.drawableBoundingBoxElement = this.DrawBoundingBox;
        this.drawer.bottomLineVisible = this.drawBottomLine;
        this.drawer.skyLineVisible = this.drawSkyLine;

        // Set page width
        let width: number = pageWidth ?? this.getContainerContentWidth();
        // A single horizontal staffline is as wide as its page. With nothing laid out, e.g. with no instrument visible, there's no
        //   page, and no backend to size.
        const firstPage: GraphicalMusicPage = this.graphic.MusicPages[0];
        if (this.rules.RenderSingleHorizontalStaffline && firstPage) {
            width = (this.EngravingRules.PageLeftMargin + firstPage.PositionAndShape.Size.width + this.EngravingRules.PageRightMargin)
                * 10 * this.zoom;
            // this.container.style.width = width + "px";
            // console.log("width: " + width)
        }
        // TODO width may need to be coordinated with render() where width is also used
        let height: number;
        const canvasDimensionsLimit: number = 32767; // browser limitation. Chrome/Firefox (16 bit, 32768 causes an error).
        // Could be calculated by canvas-size module.
        // see #678 on Github and here: https://stackoverflow.com/a/11585939/10295942

        // TODO check if resize is necessary. set needResize or something when size was changed
        for (const page of this.graphic.MusicPages) {
            if (page.PageNumber > this.rules.MaxPageToDrawNumber) {
                break; // don't add the bounding boxes of pages that aren't drawn to the container height etc
            }
            const backend: VexFlowBackend = this.createBackend(this.backendType, page);
            const sizeWarningPartTwo: string = " exceeds CanvasBackend limit of 32767. Cutting off score.";
            if (backend.getOSMDBackendType() === BackendType.Canvas && width > canvasDimensionsLimit) {
                log.warn("[OSMD] Warning: width of " + width + sizeWarningPartTwo);
                width = canvasDimensionsLimit;
            }
            if (this.rules.PageFormat && !this.rules.PageFormat.IsUndefined) {
                height = width / this.rules.PageFormat.aspectRatio;
                // console.log("pageformat given. height: " + page.PositionAndShape.Size.height);
            } else {
                height = page.PositionAndShape.Size.height;
                height += this.rules.PageBottomMargin;
                if (backend.getOSMDBackendType() === BackendType.Canvas) {
                    height += 0.1; // Canvas bug: cuts off bottom pixel with PageBottomMargin = 0. Doesn't happen with SVG.
                    //  we could only add 0.1 if PageBottomMargin === 0, but that would mean a margin of 0.1 has no effect compared to 0.
                }
                //height += this.rules.CompactMode ? this.rules.PageTopMarginNarrow : this.rules.PageTopMargin;
                // adding the PageTopMargin with a composer label leads to the margin also added to the bottom of the page
                height += page.PositionAndShape.BorderTop;
                // try to respect elements like composer cut off: this gets messy.
                // if (page.PositionAndShape.BorderTop < 0 && this.rules.PageTopMargin === 0) {
                //     height += page.PositionAndShape.BorderTop + this.rules.PageTopMargin;
                // }
                if (this.rules.RenderTitle) {
                    height += this.rules.TitleTopDistance;
                }
                height *= this.zoom * 10.0;
                // console.log("pageformat not given. height: " + page.PositionAndShape.Size.height);
            }
            if (backend.getOSMDBackendType() === BackendType.Canvas && height > canvasDimensionsLimit) {
                log.warn("[OSMD] Warning: height of " + height + sizeWarningPartTwo);
                height = Math.min(height, canvasDimensionsLimit); // this cuts off the the score, but doesn't break rendering.
                // TODO optional: reduce zoom to fit the score within the limit.
            }

            backend.resize(width, height); // this resets strokeStyle for Canvas
            backend.clear(); // set bgcolor if defined (this.rules.PageBackgroundColor, see OSMDOptions)
            backend.getContext().setFillStyle(this.rules.DefaultColorMusic);
            backend.getContext().setStrokeStyle(this.rules.DefaultColorMusic); // needs to be set after resize()
            this.drawer.Backends.push(backend);
            this.graphic.drawer = this.drawer;
        }
    }

    /**
     * Returns the width in pixels of the container's content box, which the page fills.
     * The page is drawn inside the container's border and padding, so a page as wide as the container's offsetWidth
     * overflowed the container by their width.
     * Like offsetWidth, the width includes a vertical scrollbar of the container. clientWidth doesn't: it would change
     * when createOrRefreshRenderBackend() empties the container and the scrollbar disappears, after render() used it
     * for the layout. It's also 0 in generateImages_browserless, which only sets offsetWidth.
     */
    protected getContainerContentWidth(): number {
        if (this.fixedLayoutWidth !== undefined) {
            return this.fixedLayoutWidth; // the warm keeper, which isn't in the page (see KeepWarm)
        }
        const style: CSSStyleDeclaration = window.getComputedStyle(this.container);
        // NaN -> 0: jsdom (generateImages_browserless) returns "medium" for a border that isn't set
        const pixels: (length: string) => number = (length: string): number => parseFloat(length) || 0;
        const contentWidth: number = this.container.offsetWidth - pixels(style.borderLeftWidth) - pixels(style.borderRightWidth)
            - pixels(style.paddingLeft) - pixels(style.paddingRight);
        return Math.max(0, contentWidth); // e.g. a hidden container: offsetWidth 0
    }

    /**
     * Returns the container's content width (see getContainerContentWidth()) that doesn't depend on the pages this instance
     * drew into it before: the width render() lays the sheet out for, and makes its pages as wide as.
     * A container can be as wide as its content needs, e.g. a flex item with `flex-basis: auto`, a float or an inline-block:
     * with the previous render's pages in it, its width follows theirs, not the space it has now. So the pages are taken
     * out of the layout for the measurement (display: none). Meanwhile the container keeps its height, and with it a
     * scrollbar that the pages cause, e.g. the web page's for a sheet taller than the window: the new pages will cause
     * it again.
     * Before, render() read the width with the previous pages for the layout, and createOrRefreshRenderBackend() read it
     * again without them (and without their scrollbar) for the pages: the two differed for such containers, and
     * by the scrollbar's width from the second render on.
     */
    private getContainerContentWidthWithoutPages(): number {
        const pageElements: HTMLElement[] = [];
        for (const backend of this.drawer?.Backends ?? []) {
            const element: HTMLElement = backend.getRenderElement();
            if (element?.parentElement === this.container) {
                pageElements.push(element);
            }
        }
        if (pageElements.length === 0 || this.fixedLayoutWidth !== undefined) {
            return this.getContainerContentWidth();
        }
        const containerStyle: CSSStyleDeclaration = this.container.style;
        const minHeight: string = containerStyle.minHeight;
        containerStyle.minHeight = this.container.offsetHeight + "px";
        const displays: string[] = pageElements.map((element: HTMLElement) => element.style.display);
        for (const element of pageElements) {
            element.style.display = "none";
        }
        const width: number = this.getContainerContentWidth();
        pageElements.forEach((element: HTMLElement, index: number) => {
            element.style.display = displays[index];
        });
        containerStyle.minHeight = minHeight;
        return width;
    }

    // for now SVG only, see generateImages_browserless (PNG/SVG)
    public exportSVG(): void {
        if (!this.drawer) {
            return;
        }
        for (const backend of this.drawer.Backends) {
            if (backend instanceof SvgVexFlowBackend) {
                (backend as SvgVexFlowBackend).export();
            }
            // if we add CanvasVexFlowBackend exporting, rename function to export() or exportImages() again
        }
    }

    /**
     * Export the loaded music sheet as a MIDI file.
     * @param options Optional MIDI export options
     * @returns Uint8Array containing the MIDI file data, or undefined if no sheet is loaded
     */
    public exportMIDI(options?: MidiExportOptions): Uint8Array | undefined {
        if (!this.sheet) {
            log.warn("[OSMD] exportMIDI(): No music sheet loaded.");
            return undefined;
        }
        const exporter: MidiExporter = new MidiExporter(this.sheet, options);
        return exporter.export();
    }

    /**
     * Export the loaded music sheet as a MIDI file and trigger a download.
     * @param filename Optional filename for the download (default: based on sheet title)
     * @param options Optional MIDI export options
     */
    public exportMIDIDownload(filename?: string, options?: MidiExportOptions): void {
        if (!this.sheet) {
            log.warn("[OSMD] exportMIDIDownload(): No music sheet loaded.");
            return;
        }
        const exporter: MidiExporter = new MidiExporter(this.sheet, options);
        exporter.exportAndDownload(filename);
    }

    /** States whether the render() function can be safely called. */
    public IsReadyToRender(): boolean {
        return this.graphic !== undefined;
    }

    /**
     * Gets all lyric entries across all pages and measures.
     * Useful for adding hover/click handlers to lyrics.
     * Returns empty array if sheet is not loaded or rendered.
     */
    public getAllLyricEntries(): GraphicalLyricEntry[] {
        if (!this.graphic) {
            return [];
        }
        return this.graphic.getAllLyricEntries();
    }

    /**
     * Gets all chord symbol containers across all pages and measures.
     * Useful for adding hover/click handlers to harmony symbols.
     * Returns empty array if sheet is not loaded or rendered.
     */
    public getAllChordSymbolContainers(): GraphicalChordSymbolContainer[] {
        if (!this.graphic) {
            return [];
        }
        return this.graphic.getAllChordSymbolContainers();
    }

    /**
     * Clears what OSMD has drawn on its canvas, and releases the loaded sheet, so that its memory can be freed while you keep
     * this instance: the sheet's data, its graphical objects and the SVG elements they drew.
     * To show a sheet again, call load() and render(). Until then, the cursors stay hidden, and moving them does nothing.
     * When you don't need the instance anymore, call dispose(), which removes OSMD's listeners on the window.
     */
    public clear(): void {
        this.cancelPendingLoad(); // otherwise a file that is still being fetched or unzipped would be set as the sheet later
        const sheet: MusicSheet = this.sheet;
        this.drawer?.clear(); // also releases the drawn sheet
        this.reset(); // without this, resize will draw loaded sheet again
        // reset() only removes this instance's own references to the sheet. The other ones kept the whole rendering in memory,
        //   as the graphical objects reference the SVG elements they drew.
        this.cursors.forEach(cursor => {
            cursor.releaseSheet();
        });
        this.rules.clearMusicSheetObjects(); // the graphical notes of the notes (see EngravingRules.GNote())
        // the sheet's links to its graphical measures: an app that keeps the sheet (osmd.Sheet) would keep the rendering, too
        for (const sourceMeasure of sheet?.SourceMeasures ?? []) {
            sourceMeasure.VerticalMeasureList = undefined;
        }
    }

    /**
     * Removes the listeners OSMD added to the window and the container's reference to this instance, and ends incremental
     * rendering (see renderNext()). Call it when you don't need this instance anymore but keep its container, e.g. to show
     * something else in it: with autoResize (on by default), the container keeps the instance in memory with its loaded and
     * rendered sheet, so that the drawing adapts to the window's size, and a resize would draw the sheet again, replacing the
     * container's content.
     * Without dispose(), an instance is freed together with its container, once the container was removed from the page and
     * your code references neither. Except while incremental rendering on scroll (enableIncrementalRenderingOnScroll()) hasn't
     * drawn the whole sheet: its scroll listener keeps the instance. And in browsers without WeakRef (Chrome before 84,
     * Firefox before 79, Safari before 14.1), the window's resize listener keeps it.
     * A load() that hasn't finished is aborted (see load()).
     * The drawing stays in the container, and the instance stays usable: load() and render() work, and
     * setOptions({autoResize: true}) adds the listeners again. To free the sheet's memory while you keep the instance,
     * call clear(). To remove the instance's elements from the container as well and end its use, call destroy() instead:
     * that is the call for a component's cleanup (e.g. React's useEffect cleanup, Vue's onUnmounted).
     * @example
     * osmd.dispose();
     * container.innerHTML = ""; // the instance can be garbage collected now, once your code doesn't reference it anymore
     */
    public dispose(): void {
        this.cancelPendingLoad();
        if (this.disposeResizeListener) {
            this.disposeResizeListener();
        }
        this.disableIncrementalRenderingOnScroll(); // the scroll listener of enableIncrementalRenderingOnScroll(), if on
        this.resetIncrementalRendering(); // so that a batch it scheduled (requestAnimationFrame) doesn't render anymore
        this.stopUsingWarmKeeper();
    }

    /**
     * Ends the use of this instance and leaves the container as OSMD found it: does what clear() and dispose() do (releases
     * the sheet, aborts a load() in progress, removes OSMD's listeners and observers), and removes the instance's page
     * elements and cursors from the container. The container's other children stay, e.g. elements of your app.
     * Afterwards load() is rejected and render() throws, with an Error that names destroy(): create a new instance to
     * show a score again, in the same container or another one. Calling destroy() again does nothing.
     * @example
     * // React
     * useEffect(() => {
     *     const osmd = new OpenSheetMusicDisplay(containerRef.current);
     *     osmd.load(url).then(() => osmd.render()).catch((error) => {
     *         if (error.name !== "AbortError") { // not an error: destroy() aborted the load
     *             console.error(error);
     *         }
     *     });
     *     return () => osmd.destroy();
     * }, [url]);
     */
    public destroy(): void {
        if (this.destroyed) {
            return;
        }
        this.clear();
        this.dispose();
        for (const cursor of this.cursors) {
            cursor?.Dispose(); // also removes its image
        }
        this.cursors.length = 0;
        this.removeOwnPages();
        this.destroyed = true;
    }

    /** Returns the currently committed interactive range selection, if any. */
    public getRangeSelection(): RangeSelectionPayload {
        if (!this.dragStartAnchor || !this.dragCurrentAnchor) {
            return undefined;
        }
        return this.createSelectionPayload("committed", this.dragStartAnchor, this.dragCurrentAnchor, false);
    }

    /** Programmatically sets the interactive range selection using absolute score timestamps. */
    public setRangeSelection(start: Fraction, end: Fraction): void {
        if (!this.graphic || !start || !end) {
            return;
        }
        const startAnchor: RangeSelectionAnchor = this.createAnchorFromTimestamp(start);
        const endAnchor: RangeSelectionAnchor = this.createAnchorFromTimestamp(end);
        if (!startAnchor || !endAnchor) {
            return;
        }
        this.dragStartAnchor = startAnchor;
        this.dragCurrentAnchor = endAnchor;
        this.pendingTouchRangeStartAnchor = undefined;
        this.renderRangeSelection();
        this.emitRangeSelection("committed", startAnchor, endAnchor, false);
    }

    /** Clears the interactive range selection and removes all related overlays. */
    public clearRangeSelection(emitCallback: boolean = true): void {
        const hadSelection: boolean = !!(this.dragStartAnchor && this.dragCurrentAnchor);
        const startAnchor: RangeSelectionAnchor = this.dragStartAnchor;
        const endAnchor: RangeSelectionAnchor = this.dragCurrentAnchor;
        this.dragStartAnchor = undefined;
        this.dragCurrentAnchor = undefined;
        this.pendingTouchRangeStartAnchor = undefined;
        this.activeDragBound = "both";
        this.isRangeDragging = false;
        this.maskDragChromePrepared = false;
        this.emitRangeHandleDragging(false);
        this.resetTouchGestureState();
        this.renderRangeSelection();
        if (emitCallback && hadSelection && startAnchor && endAnchor) {
            this.emitRangeSelection("cleared", startAnchor, endAnchor, false);
        }
    }

    /** Set OSMD rendering options using an IOSMDOptions object.
     *  Can be called during runtime. Also called by constructor.
     *  For example, setOptions({autoResize: false}) will disable autoResize even during runtime.
     *  An option that is given is applied, also when it turns something back on (e.g. drawFingerings: true after false);
     *  an option that is left out (or undefined) keeps its value. The options object isn't changed.
     *  Without an options object (undefined or null), this does nothing but warn.
     */
    public setOptions(options: IOSMDOptions): void {
        if (!this.rules) {
            this.rules = new EngravingRules();
            // from the start, not only from load() on, so that an app can e.g. create a GraphicalLabel before it loads a score
            this.rules.TextMeasurer = new VexFlowTextMeasurer(this.rules);
        }
        if (!this.drawingParameters && !options?.drawingParameters) {
            this.drawingParameters = new DrawingParameters(DrawingParametersEnum.default, this.rules);
            // if "default", will be created below
        } else if (options?.drawingParameters) {
            if (!this.drawingParameters) {
                this.drawingParameters = new DrawingParameters(DrawingParametersEnum[options.drawingParameters], this.rules);
            } else {
                this.drawingParameters.DrawingParametersEnum =
                    (<any>DrawingParametersEnum)[options.drawingParameters.toLowerCase()];
                    // see DrawingParameters.ts: set DrawingParametersEnum, and DrawingParameters.ts:setForCompactTightMode()
            }
        }
        if (options === undefined || options === null) {
            log.warn("warning: osmd.setOptions() called without an options parameter, has no effect."
                + "\n" + "example usage: osmd.setOptions({drawCredits: false, drawPartNames: false})");
            return;
        }
        if (options.onXMLRead) {
            this.OnXMLRead = options.onXMLRead;
        }
        if (options.rangeSelection !== undefined) {
            if (options.rangeSelection.callbacks !== undefined) {
                this.rangeSelection.callbacks = {
                    ...this.rangeSelection.callbacks,
                    ...options.rangeSelection.callbacks
                };
            }
            if (options.rangeSelection.options !== undefined) {
                this.rangeSelection.options = {
                    ...this.rangeSelection.options,
                    ...options.rangeSelection.options
                };
                if (options.rangeSelection.options.enabled !== undefined) {
                    this.rangeSelection.enabled = options.rangeSelection.options.enabled;
                }
            }
            if (options.rangeSelection.enabled !== undefined) {
                this.rangeSelection.enabled = options.rangeSelection.enabled;
            }
        }
        // Backwards compatibility with legacy top-level range options/callbacks.
        if ("onRangeSelectionChange" in options) {
            this.rangeSelection.callbacks.onChange = options.onRangeSelectionChange;
        }
        if ("onRangeSelectionLoopRequest" in options) {
            this.rangeSelection.callbacks.onLoopRequest = options.onRangeSelectionLoopRequest;
        }
        if ("onRangeSelectionClearRequest" in options) {
            this.rangeSelection.callbacks.onClearRequest = options.onRangeSelectionClearRequest;
        }
        if ("onRangeSelectionControlsRender" in options) {
            this.rangeSelection.callbacks.onControlsRender = options.onRangeSelectionControlsRender;
        }
        if ("onRangeHandleDraggingChange" in options) {
            this.rangeSelection.callbacks.onHandleDraggingChange = options.onRangeHandleDraggingChange;
        }
        if (options.interactiveRangeSelection !== undefined) {
            this.rangeSelection.enabled = options.interactiveRangeSelection;
        }
        if (options.interactiveRangeSelectionOptions !== undefined) {
            this.rangeSelection.options = {
                ...this.rangeSelection.options,
                ...options.interactiveRangeSelectionOptions
            };
            if (options.interactiveRangeSelectionOptions.enabled !== undefined) {
                this.rangeSelection.enabled = options.interactiveRangeSelectionOptions.enabled;
            }
        }
        if (options.onRender) {
            this.OnRender = options.onRender;
        }

        const backendNotInitialized: boolean = !this.drawer || !this.drawer.Backends || this.drawer.Backends.length < 1;
        let needBackendUpdate: boolean = backendNotInitialized;
        if (options.backend !== undefined) {
            const backendTypeGiven: BackendType = OSMDOptions.BackendTypeFromString(options.backend);
            needBackendUpdate = needBackendUpdate || this.backendType !== backendTypeGiven;
            this.backendType = backendTypeGiven;
        }
        this.needBackendUpdate = needBackendUpdate;
        // TODO this is a necessary step during the OSMD constructor. Maybe move this somewhere else

        // individual drawing parameters options
        if (options.autoBeam !== undefined) { // only change an option if it was given in options, otherwise it will be undefined
            this.rules.AutoBeamNotes = options.autoBeam;
        }
        if (options.autoBeamOptions) {
            const autoBeamOptions: AutoBeamOptions = { ...options.autoBeamOptions }; // a copy: the object given isn't changed
            if (autoBeamOptions.maintain_stem_directions === undefined) {
                autoBeamOptions.maintain_stem_directions = false;
            }
            this.rules.AutoBeamOptions = autoBeamOptions;
            if (autoBeamOptions.groups && autoBeamOptions.groups.length) {
                for (const fraction of autoBeamOptions.groups) {
                    if (fraction.length !== 2) {
                        throw new Error("Each fraction in autoBeamOptions.groups must be of length 2, e.g. [3,4] for beaming three fourths");
                    }
                }
            }
        }
        if (options.percussionOneLineCutoff !== undefined) {
            this.rules.PercussionOneLineCutoff = options.percussionOneLineCutoff;
        }
        if (this.rules.PercussionOneLineCutoff !== 0 &&
            options.percussionForceVoicesOneLineCutoff !== undefined) {
            this.rules.PercussionForceVoicesOneLineCutoff = options.percussionForceVoicesOneLineCutoff;
        }
        if (options.alignRests !== undefined) {
            this.rules.AlignRests = options.alignRests;
        }
        if (options.calculateMultiVoiceRestCollisions !== undefined) {
            this.rules.CalculateMultiVoiceRestCollisions = options.calculateMultiVoiceRestCollisions;
        }
        if (options.coloringMode !== undefined) {
            this.setColoringMode(options);
        }
        if (options.coloringEnabled !== undefined) {
            this.rules.ColoringEnabled = options.coloringEnabled;
        }
        if (options.colorStemsLikeNoteheads !== undefined) {
            this.rules.ColorStemsLikeNoteheads = options.colorStemsLikeNoteheads;
        }
        if (options.disableCursor !== undefined) {
            this.drawingParameters.drawCursors = !options.disableCursor; // false draws the cursors again, from the next render()
        }

        // alternative to if block: this.drawingsParameters.drawCursors = options.drawCursors !== false. No if, but always sets drawingParameters.
        // note that every option can be undefined, which doesn't mean the option should be set to false.
        if (options.drawHiddenNotes !== undefined) {
            this.drawingParameters.DrawHiddenNotes = options.drawHiddenNotes; // sets EngravingRules.RenderHiddenNotes
        }
        if (options.hiddenNotesColor !== undefined) {
            this.rules.HiddenNotesColor = options.hiddenNotesColor;
        }
        if (options.drawCredits !== undefined) {
            this.drawingParameters.DrawCredits = options.drawCredits; // sets DrawComposer, DrawTitle, DrawSubtitle, DrawLyricist.
        }
        if (options.drawComposer !== undefined) {
            this.drawingParameters.DrawComposer = options.drawComposer;
        }
        if (options.drawTitle !== undefined) {
            this.drawingParameters.DrawTitle = options.drawTitle;
        }
        if (options.drawSubtitle !== undefined) {
            this.drawingParameters.DrawSubtitle = options.drawSubtitle;
        }
        if (options.drawLyricist !== undefined) {
            this.drawingParameters.DrawLyricist = options.drawLyricist;
        }
        if (options.drawMetronomeMarks !== undefined) {
            this.rules.MetronomeMarksDrawn = options.drawMetronomeMarks;
        }
        if (options.drawSwingOnly !== undefined) {
            this.rules.DrawSwingOnly = options.drawSwingOnly;
        }
        if (options.drawDynamicTempoLabel !== undefined) {
            this.rules.DrawDynamicTempoLabel = options.drawDynamicTempoLabel;
        }
        if (options.dynamicTempoLabelBpm !== undefined) {
            this.rules.DynamicTempoLabelBpm = options.dynamicTempoLabelBpm;
        }
        if (options.drawPartNames !== undefined) {
            this.drawingParameters.DrawPartNames = options.drawPartNames; // indirectly writes to EngravingRules

            // by default, disable part abbreviations too, unless set explicitly.
            if (!options.drawPartAbbreviations) {
                this.rules.RenderPartAbbreviations = options.drawPartNames;
            }
        }
        if (options.drawPartAbbreviations !== undefined) {
            this.rules.RenderPartAbbreviations = options.drawPartAbbreviations;
        }
        if (options.drawPartAbbreviationsOnFirstSystem !== undefined) {
            this.rules.RenderPartAbbreviationsOnFirstSystem = options.drawPartAbbreviationsOnFirstSystem;
        }
        if (options.drawFingerings !== undefined) {
            this.rules.RenderFingerings = options.drawFingerings;
        }
        if (options.drawMeasureNumbers !== undefined) {
            this.rules.RenderMeasureNumbers = options.drawMeasureNumbers;
        }
        if (options.drawMeasureNumbersOnlyAtSystemStart !== undefined) {
            this.rules.RenderMeasureNumbersOnlyAtSystemStart = options.drawMeasureNumbersOnlyAtSystemStart;
        }
        if (options.drawLyrics !== undefined) {
            this.rules.RenderLyrics = options.drawLyrics;
        }
        if (options.drawTimeSignatures !== undefined) {
            this.rules.RenderTimeSignatures = options.drawTimeSignatures;
        }
        if (options.drawSlurs !== undefined) {
            this.rules.RenderSlurs = options.drawSlurs;
        }
        if (options.measureNumberInterval !== undefined) {
            this.rules.MeasureNumberLabelOffset = options.measureNumberInterval;
        }
        if (options.useXMLMeasureNumbers !== undefined) {
            this.rules.UseXMLMeasureNumbers = options.useXMLMeasureNumbers;
        }
        if (options.fingeringPosition !== undefined) {
            this.rules.FingeringPosition = AbstractExpression.PlacementEnumFromString(options.fingeringPosition);
        }
        if (options.fingeringInsideStafflines !== undefined) {
            this.rules.FingeringInsideStafflines = options.fingeringInsideStafflines;
        }
        if (options.newSystemFromXML !== undefined) {
            this.rules.NewSystemAtXMLNewSystemAttribute = options.newSystemFromXML;
        }
        if (options.newSystemFromNewPageInXML !== undefined) {
            this.rules.NewSystemAtXMLNewPageAttribute = options.newSystemFromNewPageInXML;
        }
        if (options.newPageFromXML !== undefined) {
            this.rules.NewPageAtXMLNewPageAttribute = options.newPageFromXML;
        }
        if (options.fillEmptyMeasuresWithWholeRest !== undefined) {
            this.rules.FillEmptyMeasuresWithWholeRest = options.fillEmptyMeasuresWithWholeRest;
        }
        if (options.drawRestOnlyMeasuresAsWholeMeasureRest !== undefined) {
            this.rules.DrawRestOnlyMeasuresAsWholeMeasureRest = options.drawRestOnlyMeasuresAsWholeMeasureRest;
        }
        if (options.followCursor !== undefined) {
            this.FollowCursor = options.followCursor;
        }
        if (options.setWantedStemDirectionByXml !== undefined) {
            this.rules.SetWantedStemDirectionByXml = options.setWantedStemDirectionByXml;
        }
        if (options.darkMode) {
            this.rules.applyDefaultColorMusic("#FFFFFF");
            this.rules.PageBackgroundColor = "#000000";
            this.rules.DarkModeEnabled = true;
        } else if (options.darkMode === false) { // not if undefined!
            this.rules.applyDefaultColorMusic("#000000");
            this.rules.PageBackgroundColor = undefined;
            this.rules.DarkModeEnabled = false;
        }
        if (options.defaultColorMusic) {
            this.rules.applyDefaultColorMusic(options.defaultColorMusic);
        }
        if (options.defaultColorNotehead) {
            this.rules.DefaultColorNotehead = options.defaultColorNotehead;
        }
        if (options.defaultColorRest) {
            this.rules.DefaultColorRest = options.defaultColorRest;
        }
        if (options.defaultColorStem) {
            this.rules.DefaultColorStem = options.defaultColorStem;
        }
        if (options.defaultColorLabel) {
            this.rules.DefaultColorLabel = options.defaultColorLabel;
        }
        if (options.defaultColorTitle) {
            this.rules.DefaultColorTitle = options.defaultColorTitle;
        }
        if (options.defaultFontFamily) {
            this.rules.DefaultFontFamily = options.defaultFontFamily; // default "Times New Roman", also used if font family not found
        }
        if (options.defaultFontStyle !== undefined) { // FontStyles.Regular is 0
            this.rules.DefaultFontStyle = options.defaultFontStyle; // e.g. FontStyles.Bold
        }
        if (options.drawUpToMeasureNumber >= 0) {
            this.rules.MaxMeasureToDrawIndex = Math.max(options.drawUpToMeasureNumber - 1, 0);
            this.rules.MaxMeasureToDrawNumber = options.drawUpToMeasureNumber;
        }
        if (options.drawFromMeasureNumber >= 0) {
            this.rules.MinMeasureToDrawIndex = Math.max(options.drawFromMeasureNumber - 1, 0);
            this.rules.MinMeasureToDrawNumber = options.drawFromMeasureNumber;
            // if there's a pickup measure (index and number 0), the start index might need to be + 1
            //   depending on which measure you start rendering from (measure 2 for example, instead of 0),
            //   so it is currently useful to store this option value separately from the index, to readjust the index.
        }
        if (options.drawUpToPageNumber >= 0) { // not if undefined
            this.rules.MaxPageToDrawNumber = options.drawUpToPageNumber || Number.MAX_VALUE; // 0: no limit
        }
        if (options.drawUpToSystemNumber >= 0) {
            this.rules.MaxSystemToDrawNumber = options.drawUpToSystemNumber || Number.MAX_VALUE; // 0: no limit
        }
        if (options.tupletsRatioed !== undefined) {
            this.rules.TupletsRatioed = options.tupletsRatioed;
        }
        if (options.tupletsBracketed !== undefined) {
            this.rules.TupletsBracketed = options.tupletsBracketed;
        }
        if (options.tripletsBracketed !== undefined) {
            this.rules.TripletsBracketed = options.tripletsBracketed;
        }
        if (options.keepWarm !== undefined) {
            this.KeepWarm = options.keepWarm;
        }
        if (options.autoResize) {
            this.AutoResizeEnabled = true;
        } else if (options.autoResize === false) { // not undefined
            this.AutoResizeEnabled = false; // also removes the window's resize listener, see there
        }
        if (options.pageFormat !== undefined) { // only change this option if it was given, see above
            this.setPageFormat(options.pageFormat);
        }
        if (options.pageBackgroundColor !== undefined) {
            this.rules.PageBackgroundColor = options.pageBackgroundColor;
        }
        if (options.renderSingleHorizontalStaffline !== undefined) {
            this.rules.RenderSingleHorizontalStaffline = options.renderSingleHorizontalStaffline;
        }
        if (options.spacingFactorSoftmax !== undefined) {
            this.rules.SoftmaxFactorVexFlow = options.spacingFactorSoftmax;
        }
        if (options.rhythmicSpacingRatio !== undefined) {
            this.rules.RhythmicSpacingRatio = options.rhythmicSpacingRatio;
        }
        if (options.spacingBetweenTextLines !== undefined) {
            this.rules.SpacingBetweenTextLines = options.spacingBetweenTextLines;
        }
        if (options.stretchLastSystemLine !== undefined) {
            this.rules.StretchLastSystemLine = options.stretchLastSystemLine;
        }
        if (options.autoGenerateMultipleRestMeasuresFromRestMeasures !== undefined) {
            this.rules.AutoGenerateMultipleRestMeasuresFromRestMeasures = options.autoGenerateMultipleRestMeasuresFromRestMeasures;
        }
        if (options.cursorsOptions !== undefined) {
            this.cursorsOptions = options.cursorsOptions;
        } else if (!this.cursorsOptions) {
            // the standard cursor, in the constructor. Later calls keep the cursors, like the other options that are left out.
            this.cursorsOptions = [{
                type: CursorType.Standard,
                color: this.EngravingRules.DefaultColorCursor,
                alpha: 0.5,
                follow: true,
                followCursorPolyfill: options.followCursorPolyfill,
                followCursorPolyfillOffsetY: options.followCursorPolyfillOffsetY,
            }];
        }
        for (const cursorOptions of this.cursorsOptions) {
            if (options.followCursorPolyfill !== undefined) {
                cursorOptions.followCursorPolyfill = options.followCursorPolyfill;
            }
            if (options.followCursorPolyfillOffsetY !== undefined) {
                cursorOptions.followCursorPolyfillOffsetY = options.followCursorPolyfillOffsetY;
            }
        }
        if (options.useGeometricSkyBottomLineCalculation !== undefined) {
            this.rules.UseGeometricSkyBottomLineCalculation = options.useGeometricSkyBottomLineCalculation;
        }
        if (options.skyBottomLineBatchMinMeasures !== undefined) {
            this.rules.SkyBottomLineBatchMinMeasures = options.skyBottomLineBatchMinMeasures;
        }
        this.syncInteractiveRangeSelection();
    }

    public setColoringMode(options: IOSMDOptions): void {
        if (options.coloringMode === ColoringModes.XML) {
            this.rules.ColoringMode = ColoringModes.XML;
            return;
        }
        const noteIndices: NoteEnum[] = [NoteEnum.C, NoteEnum.D, NoteEnum.E, NoteEnum.F, NoteEnum.G, NoteEnum.A, NoteEnum.B];
        let colorSetString: string[];
        if (options.coloringMode === ColoringModes.CustomColorSet) {
            if (!options.coloringSetCustom || options.coloringSetCustom.length !== 8) {
                throw new Error("Invalid amount of colors: With coloringModes.customColorSet, " +
                    "you have to provide a coloringSetCustom parameter (array) with 8 strings (C to B, rest note).");
            }
            // validate strings input
            for (const colorString of options.coloringSetCustom) {
                const regExp: RegExp = /^#[0-9a-fA-F]{6}$/;
                if (!regExp.test(colorString)) {
                    throw new Error(
                        "One of the color strings in options.coloringSetCustom was not a valid HTML Hex color:\n" + colorString);
                }
            }
            colorSetString = options.coloringSetCustom;
        } else if (options.coloringMode === ColoringModes.AutoColoring) {
            colorSetString = [];
            const keys: string[] = Object.keys(AutoColorSet);
            for (let i: number = 0; i < keys.length; i++) {
                colorSetString.push(AutoColorSet[keys[i]]);
            }
        } // for both cases:
        const coloringSetCurrent: Dictionary<NoteEnum | number, string> = new Dictionary<NoteEnum | number, string>();
        for (let i: number = 0; i < noteIndices.length; i++) {
            coloringSetCurrent.setValue(noteIndices[i], colorSetString[i]);
        }
        // index 7. Unfortunately -1 is not a NoteEnum value, so we can't put it into noteIndices
        coloringSetCurrent.setValue(-1, CollectionUtil.getLastElement(colorSetString));
        this.rules.ColoringSetCurrent = coloringSetCurrent;
        this.rules.ColoringMode = options.coloringMode;
    }

    /**
     * Sets the logging level of OSMD: for all OSMD instances on the page, not just this one.
     * It sets the level of OSMD's own loglevel logger (named "OSMD"), not of loglevel's root logger, so the level of the
     * page's other loglevel users stays as it is, and it isn't saved in localStorage (it lasts until the page reloads).
     *
     * Until it's called, OSMD logs with the level the root logger has when OSMD loads: `warn` by default.
     *
     * @param level `trace`, `debug`, `info`, `warn`, `error` or `silent`. Any other value logs a warning and sets `warn`.
     */
    public setLogLevel(level: string): void {
        switch (level) {
            case "trace":
                log.setLevel(log.levels.TRACE, false);
                break;
            case "debug":
                log.setLevel(log.levels.DEBUG, false);
                break;
            case "info":
                log.setLevel(log.levels.INFO, false);
                break;
            case "warn":
                log.setLevel(log.levels.WARN, false);
                break;
            case "error":
                log.setLevel(log.levels.ERROR, false);
                break;
            case "silent":
                log.setLevel(log.levels.SILENT, false);
                break;
            default:
                log.warn(`Could not set log level to ${level}. Using warn instead.`);
                log.setLevel(log.levels.WARN, false);
                break;
        }
    }

    /**
     * Returns OSMD's logging level (for all OSMD instances on the page), as a number of loglevel's `levels`:
     * 0 = trace, 1 = debug, 2 = info, 3 = warn, 4 = error, 5 = silent. See {@link setLogLevel}.
     */
    public getLogLevel(): number {
        return log.getLevel();
    }

    /**
     * Releases the loaded sheet and its graphic, and ends incremental rendering: for load(), before it reads the new sheet,
     * and for clear(). The settings stay: the options, the EngravingRules and the zoom.
     */
    protected reset(): void {
        this.systemVirtualization.invalidate();
        this.resetIncrementalRendering(); // abandon any incremental session
        if (this.drawingParameters.drawCursors) {
            this.cursors.forEach(cursor => {
                cursor.hide();
            });
        }
        this.sheet = undefined;
        this.graphic = undefined;
        this.graphicNeedsPreparation = false;
        this.rules.RenderCount = 0;
        this.staffOpacityOverrides.clear();
        this.invalidateReadAheadStaffEntryIndex();
        this.rangeSelectionElementCollector.invalidate();
        this.clearRangeSelection(false);
        this.hoverAnchor = undefined;
        this.detachRangeSelectionListeners();
        this.removeRangeSelectionOverlay();
        this.layoutContainerWidth = undefined; // no layout: the container's observer doesn't render a sheet before the app does
    }

    /**
     * Attach the appropriate handler for autoResize: an observer of the container's size, or a listener to the window.onResize
     * event (see handleResize())
     */
    protected autoResize(): void {
        if (this.destroyed) {
            return;
        }
        const self: OpenSheetMusicDisplay = this;
        this.handleResize(
            () => {
                // empty
            },
            () => {
                // The following code is probably not needed
                // (the width should adapt itself to the max allowed)
                //let width: number = Math.max(
                //    document.documentElement.clientWidth,
                //    document.body.scrollWidth,
                //    document.documentElement.scrollWidth,
                //    document.body.offsetWidth,
                //    document.documentElement.offsetWidth
                //);
                //self.container.style.width = width + "px";

                // recalculate beams, are otherwise not updated and can detach from stems, see #724
                if (this.graphic?.GetCalculator instanceof VexFlowMusicSheetCalculator) { // null and type check
                    (this.graphic.GetCalculator as VexFlowMusicSheetCalculator).beamsNeedUpdate = true;
                }
                if (self.IsReadyToRender()) {
                    self.renderAndScrollBack(); // just calling render() will scroll to the top of the page
                }
            }
        );
    }

    /** Re-render and scroll back to previous scroll bar y position in percent.
     * If the document keeps the same height/length, the scroll bar position will basically be unchanged.
     * For example, if you scroll to the bottom of the page, resize by one pixel (or enable dark mode) and call this,
     *   for the human eye there will be no detectable scrolling or change in the scroll position at all.
     * If you just call render() instead of renderAndScrollBack(),
     *   it will scroll you back to the top of the page, even if you were scrolled to the bottom before. */
    public renderAndScrollBack(): void {
        const previousScrollY: number = window.scrollY;
        const previousScrollHeight: number = document.body.scrollHeight; // height of page
        const previousScrollYPercent: number = previousScrollY / previousScrollHeight;
        this.render();
        const newScrollHeight: number = document.body.scrollHeight; // height of page
        const newScrollY: number = newScrollHeight * previousScrollYPercent;
        window.scrollTo({
            top: newScrollY,
            behavior: "instant" // visually, there is no change in the scroll bar position, as it's the same as before.
        });
    }

    public getActiveOctaveShift(timestamp: Fraction, staffIndex: number, measureIndex: number): OctaveShift | undefined {
        if (!this.sheet || !this.sheet.SourceMeasures?.[measureIndex]) {
            return undefined;
        }
        const measure: MusicSheet["SourceMeasures"][number] = this.sheet.SourceMeasures[measureIndex];
        const measureStart: Fraction = measure.AbsoluteTimestamp;
        const measureEnd: Fraction = Fraction.plus(measureStart, measure.Duration);
        let absTs: Fraction = timestamp;
        if (timestamp.lt(measureStart) || timestamp.gt(measureEnd)) {
            absTs = Fraction.plus(measureStart, timestamp);
        }
        const sourceMeasures: MusicSheet["SourceMeasures"] = this.sheet.SourceMeasures;
        const lastSourceMeasure: MusicSheet["SourceMeasures"][number] = sourceMeasures[sourceMeasures.length - 1];
        const sheetEndTs: Fraction = Fraction.plus(lastSourceMeasure.AbsoluteTimestamp, lastSourceMeasure.Duration);
        for (let m: number = 0; m < sourceMeasures.length; m++) {
            const sm: MusicSheet["SourceMeasures"][number] = sourceMeasures[m];
            const expressions: MultiExpression[] = sm.StaffLinkedExpressions?.[staffIndex];
            if (!expressions) {
                continue;
            }
            for (let i: number = 0; i < expressions.length; i++) {
                const multi: MultiExpression = expressions[i];
                const shift: OctaveShift = multi.OctaveShiftStart || multi.OctaveShiftEnd;
                if (!shift) {
                    continue;
                }
                const start: Fraction = shift.ParentStartMultiExpression?.AbsoluteTimestamp;
                const end: Fraction = shift.ParentEndMultiExpression?.AbsoluteTimestamp ?? sheetEndTs;
                if (start && start.lte(absTs) && !end.lt(absTs)) {
                    return shift;
                }
            }
        }
        return undefined;
    }

    public hasActiveOctaveShift(timestamp: Fraction, staffIndex: number, measureIndex: number): boolean {
        return this.getActiveOctaveShift(timestamp, staffIndex, measureIndex) !== undefined;
    }

    public getBPMTempoFromTimestamp(timestamp: Fraction, measureIndex: number): number | undefined {
        if (!this.sheet || !this.sheet.SourceMeasures?.[measureIndex]) {
            return undefined;
        }
        const measure: MusicSheet["SourceMeasures"][number] = this.sheet.SourceMeasures[measureIndex];
        const measureStart: Fraction = measure.AbsoluteTimestamp;
        const measureEnd: Fraction = Fraction.plus(measureStart, measure.Duration);
        let absTs: Fraction = timestamp;
        if (timestamp.lt(measureStart) || timestamp.gt(measureEnd)) {
            absTs = Fraction.plus(measureStart, timestamp);
        }

        const tempoExpressions: MultiTempoExpression[] = this.sheet.TimestampSortedTempoExpressionsList;
        let activeTempoExpression: MultiTempoExpression | undefined;

        for (let i: number = tempoExpressions.length - 1; i >= 0; i--) {
            const tempoExpr: MultiTempoExpression = tempoExpressions[i];
            const exprStart: Fraction = tempoExpr.AbsoluteTimestamp;

            if (exprStart.gt(absTs)) {
                continue;
            }

            if (tempoExpr.InstantaneousTempo) {
                activeTempoExpression = tempoExpr;
                break;
            }

            if (tempoExpr.ContinuousTempo) {
                const exprEnd: Fraction = tempoExpr.ContinuousTempo.AbsoluteEndTimestamp;
                if (absTs.lte(exprEnd)) {
                    activeTempoExpression = tempoExpr;
                    break;
                }
            }
        }

        if (activeTempoExpression) {
            if (activeTempoExpression.InstantaneousTempo) {
                return activeTempoExpression.InstantaneousTempo.TempoInBpm;
            }
            if (activeTempoExpression.ContinuousTempo) {
                const interpolatedTempo: number = activeTempoExpression.ContinuousTempo.getInterpolatedTempo(absTs);
                if (interpolatedTempo > 0) {
                    return interpolatedTempo;
                }
            }
        }

        if (measure.TempoInBPM > 0) {
            return measure.TempoInBPM;
        }

        return undefined;
    }

    public updateTempo(newInitialBPM: number, render: boolean = true): void {
        if (!this.sheet) {
            return;
        }

        let oldInitialTempo: number = this.sheet.getExpressionsStartTempoInBPM();
        if (oldInitialTempo === 0) {
            if (this.sheet.SourceMeasures.length > 0 && this.sheet.SourceMeasures[0].TempoInBPM > 0) {
                oldInitialTempo = this.sheet.SourceMeasures[0].TempoInBPM;
            } else {
                return;
            }
        }

        if (oldInitialTempo === newInitialBPM) {
            return;
        }

        const ratio: number = newInitialBPM / oldInitialTempo;

        for (const tempoExpr of this.sheet.TimestampSortedTempoExpressionsList) {
            if (tempoExpr.InstantaneousTempo) {
                tempoExpr.InstantaneousTempo.TempoInBpm *= ratio;
            }
            if (tempoExpr.ContinuousTempo) {
                tempoExpr.ContinuousTempo.StartTempo *= ratio;
                tempoExpr.ContinuousTempo.EndTempo *= ratio;
            }
        }

        for (const measure of this.sheet.SourceMeasures) {
            if (measure.TempoInBPM > 0) {
                measure.TempoInBPM *= ratio;
            }
        }

        if (this.sheet.TimestampSortedTempoExpressionsList.length === 0) {
            this.sheet.userStartTempoInBPM *= ratio;
        }

        if (render && this.graphic && this.drawer) {
            this.updateMetronomeMarksInSVG(ratio);
        } else if (render) {
            this.render();
        }
    }

    private updateMetronomeMarksInSVG(ratio: number): void {
        if (!this.drawer || !this.drawer.Backends) {
            return;
        }

        const backends: VexFlowBackend[] = this.drawer.Backends;
        for (const backend of backends) {
            const renderElement: HTMLElement = backend.getRenderElement?.();
            if (renderElement) {
                const bpmGroups: NodeListOf<Element> = renderElement.querySelectorAll("g.vf-bpm");
                if (bpmGroups) {
                    for (const bpmGroup of bpmGroups) {
                        const textNodes: NodeListOf<Element> = bpmGroup.querySelectorAll("text");
                        if (textNodes) {
                            for (const textNode of textNodes) {
                                const textContent: string = textNode.textContent || "";
                                const match: RegExpMatchArray | null = textContent.match(/ = (\d+)/);
                                if (match && match[1]) {
                                    const oldBpm: number = parseFloat(match[1]);
                                    const newBpm: number = Math.round(oldBpm * ratio);
                                    textNode.textContent = textContent.replace(/ = \d+/, " = " + newBpm);
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    /**
     * Updates the BPM-driven dynamic tempo label (e.g. "Andante") in-place in the SVG, without
     * re-rendering the sheet. Only has an effect when the score was rendered with the
     * drawDynamicTempoLabel option. The BPM is stored so the label is reproduced on the next render
     * (e.g. after a resize/relayout).
     * @param bpm The current quarter-note BPM.
     */
    public setDynamicTempoLabel(bpm: number): void {
        this.rules.DynamicTempoLabelBpm = bpm;
        if (!this.rules.DrawDynamicTempoLabel) {
            return;
        }
        const label: string = tempoLabelFromBpm(bpm) ?? "";
        this.updateDynamicTempoLabelInSVG(label);
    }

    private updateDynamicTempoLabelInSVG(label: string): void {
        if (!this.drawer || !this.drawer.Backends) {
            return;
        }

        const backends: VexFlowBackend[] = this.drawer.Backends;
        for (const backend of backends) {
            const renderElement: HTMLElement = backend.getRenderElement?.();
            if (!renderElement) {
                continue;
            }
            const dynamicTempoGroups: NodeListOf<Element> = renderElement.querySelectorAll("g.vf-dynamic-tempo");
            for (const group of dynamicTempoGroups) {
                const textNodes: NodeListOf<Element> = group.querySelectorAll("text");
                for (const textNode of textNodes) {
                    textNode.textContent = label;
                }
            }
        }
    }

    /**
     * Whether the sheet needs a new layout because the container's content width isn't the one of the last layout anymore
     * (see layoutContainerWidth): what the container's observer of autoResize renders again for.
     * Not before the first render() (or the start of incremental rendering): the app renders a loaded sheet when it
     * chooses to. And not while the container has no width, e.g. while it's hidden: render() would skip that; when it's
     * shown again at the width of the layout, nothing is to do.
     */
    private containerWidthNeedsLayout(): boolean {
        if (this.layoutContainerWidth === undefined || !this.graphic) {
            return false;
        }
        const width: number = this.getContainerContentWidth();
        return width > 0 && width !== this.layoutContainerWidth && width !== this.containerWidthWithPages;
    }

    /**
     * Helper function for managing the resize events of autoResize: those of a ResizeObserver on the container, or the
     * window's onResize events in browsers without ResizeObserver (see addResizeListener()).
     * The window keeps its listeners as long as the page is open, so the listener reaches this instance only through a WeakRef:
     * it doesn't keep an instance in memory, with its sheet and drawing, that the app has dropped. Instead, the container keeps
     * the instance until the listener is removed (e.g. by dispose()), so that the drawing adapts to the window's size as long as
     * the container exists, also when the app keeps no reference to the instance, e.g. after
     * `const osmd = new OpenSheetMusicDisplay(div); osmd.load(url).then(() => osmd.render());` in a function.
     * In browsers without WeakRef, the listener keeps the instance.
     * @param startCallback is the function called when resizing starts
     * @param endCallback is the function called when resizing (kind-of) ends
     */
    protected handleResize(startCallback: () => void, endCallback: () => void): void {
        this.resizeStartCallback = startCallback;
        this.resizeEndCallback = endCallback;
        const instanceRef: { deref(): OpenSheetMusicDisplay } = typeof WeakRef === "function" ?
            new WeakRef(this) : { deref: (): OpenSheetMusicDisplay => this };
        const container: HTMLElement = this.container;
        const removeListener: () => void = OpenSheetMusicDisplay.addResizeListener(instanceRef, container);
        const key: symbol = OpenSheetMusicDisplay.containerInstanceKey;
        // not enumerable. Replaces the reference to an older instance made for the same container.
        Object.defineProperty(container, key, { value: this, configurable: true });
        this.disposeResizeListener = (): void => {
            removeListener();
            if ((<any>container)[key] === this) { // not a newer instance's on the same container
                delete (<any>container)[key];
            }
            this.disposeResizeListener = undefined;
            this.resizeHandlerAttached = false; // setOptions({autoResize: true}) adds it again
        };
        this.resizeHandlerAttached = true;

        // If the sheet was already rendered when autoResize was enabled, adapt it to the
        // current width right away. (Only in that case: otherwise - e.g. with the default
        // autoResize on creation - these initial timers would re-render the first sheet the
        // application loads and renders a second time, once the timers fire.)
        // And only if the container's width changed since the layout, e.g. while autoResize was off: turning it off and on
        // again doesn't re-render a sheet that already fits.
        const renderedBeforeAttach: boolean = this.rules.RenderCount > 0;
        if (renderedBeforeAttach && this.getContainerContentWidth() !== this.layoutContainerWidth) {
            window.setTimeout(startCallback, 0);
            window.setTimeout(endCallback, 1);
        }
    }

    /**
     * Adds a listener that calls the instance's resize callbacks (see handleResize()): the start callback when a resize
     * starts, the end callback once no resize event came for 200 ms.
     * With ResizeObserver in the browser, the resize events are the changes of the container's width
     * that need a new layout (see containerWidthNeedsLayout()), whatever their cause: the window, a sidebar, a split pane,
     * a hidden container that is shown. Changes of its height don't count: each render changes it. If the width is back at
     * the layout's after the 200 ms, the end callback isn't called.
     * Otherwise they are the window's resize events, each of which ends in the end callback.
     * The listener and its timers reach the instance only through `instanceRef`. When it's a WeakRef and the instance has been
     * freed, the listener removes itself at the next resize event.
     * It's made here, not in handleResize(), because all closures made in one function call share the variables that any of
     * them uses: next to the closures there that reference `this`, the listener would keep the instance in memory too.
     * @param instanceRef a WeakRef to the instance, or (without WeakRef) an object that keeps it
     * @param container the instance's container, which is observed
     * @returns a function that removes the listener
     */
    private static addResizeListener(instanceRef: { deref(): OpenSheetMusicDisplay }, container: HTMLElement): () => void {
        let rtime: number;
        let timeout: number = undefined;
        const delta: number = 200;
        let observer: ResizeObserver;

        function containerResized(): void {
            const osmd: OpenSheetMusicDisplay = instanceRef.deref();
            if (!osmd) {
                removeListener(); // the instance was freed
            } else if (osmd.containerWidthNeedsLayout()) {
                resizeStart();
            }
        }

        function resizeStart(): void {
            const osmd: OpenSheetMusicDisplay = instanceRef.deref();
            if (!osmd) {
                removeListener(); // the instance was freed
                return;
            }
            if (!osmd.AutoResizeEnabled) {
                return;
            }
            rtime = (new Date()).getTime();
            if (!timeout) {
                osmd.resizeStartCallback();
                rtime = (new Date()).getTime();
                timeout = window.setTimeout(resizeEnd, delta);
            }
        }

        function resizeEnd(): void {
            timeout = undefined;
            window.clearTimeout(timeout);
            const osmd: OpenSheetMusicDisplay = instanceRef.deref();
            if (!osmd) {
                removeListener(); // the instance was freed
            } else if ((new Date()).getTime() - rtime < delta) {
                timeout = window.setTimeout(resizeEnd, delta);
            } else if (!observer || osmd.containerWidthNeedsLayout()) {
                osmd.resizeEndCallback();
            }
        }

        function removeListener(): void {
            if (observer) {
                observer.disconnect();
            } else if ((<any>window).detachEvent) {
                // Support IE<9
                (<any>window).detachEvent("onresize", resizeStart);
            } else {
                window.removeEventListener("resize", resizeStart);
            }
        }

        if (typeof ResizeObserver === "function") {
            observer = new ResizeObserver(containerResized);
            observer.observe(container);
        } else if ((<any>window).attachEvent) {
            // Support IE<9
            (<any>window).attachEvent("onresize", resizeStart);
        } else {
            window.addEventListener("resize", resizeStart);
        }
        return removeListener;
    }

    /** Enable or disable (hide) the cursor.
     * @param enable whether to enable (true) or disable (false) the cursor
     */
    public enableOrDisableCursors(enable: boolean): void {
        this.drawingParameters.drawCursors = enable;
        if (enable) {
            for (let i: number = 0; i < this.cursorsOptions.length; i++){
                // save previous cursor state
                const hidden: boolean = this.cursors[i]?.Hidden ?? true;
                const previousIterator: MusicPartManagerIterator = this.cursors[i]?.Iterator;
                this.cursors[i]?.hide();

                // check which page/backend to draw the cursor on (the pages may have changed since last cursor)
                let backendToDrawOn: VexFlowBackend = this.drawer?.Backends[0];
                if (backendToDrawOn && this.rules.RestoreCursorAfterRerender && this.cursors[i]) {
                    const newPageNumber: number = this.cursors[i].updateCurrentPage();
                    backendToDrawOn = this.drawer.Backends[newPageNumber - 1];
                }
                // create new cursor
                if (backendToDrawOn && backendToDrawOn.getRenderElement()) {
                    const previousCursor: Cursor = this.cursors[i];
                    this.cursors[i] = new Cursor(backendToDrawOn.getRenderElement(), this, this.cursorsOptions[i]);
                    if (previousCursor) {
                        this.cursors[i].useImageOf(previousCursor); // usually the same image: no need to draw it again
                    }
                }
                if (this.sheet && this.graphic && this.cursors[i]) { // else init is called in load()
                    this.cursors[i].init(this.sheet.MusicPartManager, this.graphic);
                }

                // restore old cursor state, if there's a cursor: it's created on a drawn page, and with no instrument visible,
                //   render() draws none (#1892)
                if (this.rules.RestoreCursorAfterRerender && this.cursors[i]) {
                    this.cursors[i].hidden = hidden;
                    if (previousIterator) {
                        this.cursors[i].iterator = previousIterator;
                        this.cursors[i].update();
                    }
                }
            }
            // remove the cursors whose options were removed, e.g. by setOptions() with fewer cursorsOptions
            for (const removedCursor of this.cursors.splice(this.cursorsOptions.length)) {
                removedCursor?.Dispose(); // also removes its image
            }
        } else { // disable cursor
            this.cursors.forEach(cursor => {
                cursor.hide();
            });
            // this.cursor = undefined;
            // TODO cursor should be disabled, not just hidden. otherwise user can just call osmd.cursor.hide().
            // however, this could cause null calls (cursor.next() etc), maybe that needs some solution.
        }
    }

    public createBackend(type: BackendType, page: GraphicalMusicPage, idOverride?: string): VexFlowBackend {
        let backend: VexFlowBackend;
        if (type === undefined || type === BackendType.SVG) {
            backend = new SvgVexFlowBackend(this.rules);
        } else {
            backend = new CanvasVexFlowBackend(this.rules);
        }
        backend.graphicalMusicPage = page; // the page the backend renders on. needed to identify DOM element to extract image/SVG
        backend.initialize(this.container, this.zoom, idOverride);
        //backend.getContext().setFillStyle(this.rules.DefaultColorMusic);
        //backend.getContext().setStrokeStyle(this.rules.DefaultColorMusic);
        // color needs to be set after resize() for CanvasBackend
        return backend;
    }

    /** Standard page format options like A4 or Letter, in portrait and landscape. E.g. PageFormatStandards["A4_P"] or PageFormatStandards["Letter_L"].
     *  One table for all OSMD instances on the page: a format added here, e.g. PageFormatStandards["B5_P"] = new PageFormat(176, 250, "B5_P"),
     *  can be used by all of them (setOptions({pageFormat: "B5_P"})). The table itself can't be replaced (readonly). */
    public static readonly PageFormatStandards: { [type: string]: PageFormat } = {
        "A3_L": new PageFormat(420, 297, "A3_L"), // id strings should use underscores instead of white spaces to facilitate use as URL parameters.
        "A3_P": new PageFormat(297, 420, "A3_P"),
        "A4_L": new PageFormat(297, 210, "A4_L"),
        "A4_P": new PageFormat(210, 297, "A4_P"),
        "A5_L": new PageFormat(210, 148, "A5_L"),
        "A5_P": new PageFormat(148, 210, "A5_P"),
        "A6_L": new PageFormat(148, 105, "A6_L"),
        "A6_P": new PageFormat(105, 148, "A6_P"),
        "Endless": PageFormat.UndefinedPageFormat,
        "Letter_L": new PageFormat(279.4, 215.9, "Letter_L"),
        "Letter_P": new PageFormat(215.9, 279.4, "Letter_P")
    };

    public static StringToPageFormat(pageFormatString: string): PageFormat {
        let pageFormat: PageFormat = PageFormat.UndefinedPageFormat; // default: 'endless' page height, take canvas/container width

        // check for widthxheight parameter, e.g. "800x600"
        if (pageFormatString.match("^[0-9]+x[0-9]+$")) {
            const widthAndHeight: string[] = pageFormatString.split("x");
            const width: number = Number.parseInt(widthAndHeight[0], 10);
            const height: number = Number.parseInt(widthAndHeight[1], 10);
            if (width > 0 && width < 32768 && height > 0 && height < 32768) {
                pageFormat = new PageFormat(width, height, `customPageFormat${pageFormatString}`);
            }
        }

        // check for formatId from OpenSheetMusicDisplay.PageFormatStandards
        pageFormatString = pageFormatString.replace(" ", "_");
        pageFormatString = pageFormatString.replace("Landscape", "L");
        pageFormatString = pageFormatString.replace("Portrait", "P");
        //console.log("change format to: " + formatId);
        if (OpenSheetMusicDisplay.PageFormatStandards.hasOwnProperty(pageFormatString)) {
            pageFormat = OpenSheetMusicDisplay.PageFormatStandards[pageFormatString];
            return pageFormat;
        }
        return pageFormat;
    }

    /** Sets page format by string. Used by setOptions({pageFormat: "A4_P"}) for example. */
    public setPageFormat(formatId: string): void {
        const newPageFormat: PageFormat = OpenSheetMusicDisplay.StringToPageFormat(formatId);
        this.needBackendUpdate = !(newPageFormat.Equals(this.rules.PageFormat));
        this.rules.PageFormat = newPageFormat;
    }

    public setCustomPageFormat(width: number, height: number): void {
        if (width > 0 && height > 0) {
            const f: PageFormat = new PageFormat(width, height);
            this.rules.PageFormat = f;
        }
    }

    //#region GETTER / SETTER
    public set DrawSkyLine(value: boolean) {
        this.drawSkyLine = value;
        if (this.drawer) {
            this.drawer.skyLineVisible = value;
            // this.render(); // note: we probably shouldn't automatically render when someone sets the setter
            //   this can cause a lot of rendering time.
        }
    }
    public get DrawSkyLine(): boolean {
        return this.drawer.skyLineVisible;
    }

    public set DrawBottomLine(value: boolean) {
        this.drawBottomLine = value;
        if (this.drawer) {
            this.drawer.bottomLineVisible = value;
            // this.render(); // note: we probably shouldn't automatically render when someone sets the setter
            //   this can cause a lot of rendering time.
        }
    }
    public get DrawBottomLine(): boolean {
        return this.drawer.bottomLineVisible;
    }
    public set DrawBoundingBox(value: string) {
        this.setDrawBoundingBox(value, true);
    }
    public get DrawBoundingBox(): string {
        return this.drawBoundingBox;
    }
    public setDrawBoundingBox(value: string, render: boolean = false): void {
        this.drawBoundingBox = value;
        if (this.drawer) {
            this.drawer.drawableBoundingBoxElement = value; // drawer is sometimes created anew, losing this value, so it's saved in OSMD now.
        }
        if (render) {
            this.renderAndScrollBack(); // may create new Drawer.
        }
    }

    public get AutoResizeEnabled(): boolean {
        return this.autoResizeEnabled;
    }
    /** Turns autoResize on or off, also during runtime, like setOptions({autoResize}).
     *  Off removes the container's observer (or the window's resize listener) and the container's reference to this
     *  instance (see handleResize()); on adds them again.
     */
    public set AutoResizeEnabled(value: boolean) {
        if (value && !this.resizeHandlerAttached) {
            this.autoResize();
        } else if (!value && this.disposeResizeListener) {
            this.disposeResizeListener();
        }
        this.autoResizeEnabled = value;
    }

    /** The scale of the drawing: 1.0 by default, e.g. 0.5 draws the sheet at half the size, with more measures per system.
     *  Call render() after setting it. Like the options, it stays set when another score is loaded (load()) and over
     *  clear(), so it can also be set before the first load(). (Up to OSMD 2.3.1, load() and clear() set it back to 1.0.)
     */
    public get Zoom(): number {
        return this.zoom;
    }
    public set Zoom(value: number) {
        this.zoom = value;
        this.zoomUpdated = true;
        if (this.graphic?.GetCalculator instanceof VexFlowMusicSheetCalculator) { // null and type check
            (this.graphic.GetCalculator as VexFlowMusicSheetCalculator).beamsNeedUpdate = this.zoomUpdated;
        }
    }

    public set FollowCursor(value: boolean) {
        this.followCursor = value;
    }

    public get FollowCursor(): boolean {
        return this.followCursor;
    }

    /** Sets this instance's calculator for transposing (osmd.Sheet.Transpose, Instrument.Transpose),
     *  e.g. new TransposeCalculator() or new IntervalTransposeCalculator(). Other OSMD instances on the page keep their own.
     *  Undefined (the default): no transposition. Stored in EngravingRules.TransposeCalculator.
     */
    public set TransposeCalculator(calculator: ITransposeCalculator) {
        if (calculator) {
            // give the calculator access to the transposition rules, e.g. EngravingRules.StrictTransposeSpelling
            calculator.rules = this.rules;
        }
        this.rules.TransposeCalculator = calculator;
    }

    /** This instance's calculator for transposing, see the setter. */
    public get TransposeCalculator(): ITransposeCalculator {
        return this.rules.TransposeCalculator;
    }

    /**
     * Read-only: does not set `Sheet.Transpose` or render. Counts printed notes that would fall outside
     * treble D4–G5 or bass F2–B3 if the score were hypothetically transposed by `transposeHalftones` (semitones).
     * Useful to know whether we want to transpose upwards or downwards, i.e. which result would be easier to read!
     */
    public countLedgerLineNotesForTransposition(transposeHalftones: number): number {
        return countLedgerLineNotesOnMusicSheet(this.sheet, transposeHalftones);
    }

    public get Sheet(): MusicSheet {
        return this.sheet;
    }
    public get Drawer(): VexFlowMusicSheetDrawer {
        return this.drawer;
    }
    public get GraphicSheet(): GraphicalMusicSheet {
        this.ensureGraphicPrepared();
        return this.graphic;
    }

    /**
     * Graphical measures are created on first use rather than in load(): the application may still change
     * instrument visibility, drawing range or rules before rendering, and preparation depends on them.
     */
    private ensureGraphicPrepared(): void {
        if (this.graphicNeedsPreparation && this.graphic) {
            this.prepareGraphic();
        }
    }

    private prepareGraphic(): void {
        this.graphic.Initialize();
        this.graphic.GetCalculator.prepareGraphicalMusicSheet();
        this.graphicNeedsPreparation = false;
        this.lastMinMeasureToDrawIndex = this.rules.MinMeasureToDrawIndex;
        this.lastMaxMeasureToDrawIndex = this.rules.MaxMeasureToDrawIndex;
    }
    /** Changes whenever the graphical score is rebuilt; 0 before a score is loaded. */
    public get LayoutGeneration(): number {
        return this.graphic?.LayoutGeneration ?? 0;
    }
    public get DrawingParameters(): DrawingParameters {
        return this.drawingParameters;
    }
    public get EngravingRules(): EngravingRules { // custom getter, useful for engraving parameter setting in Demo
        return this.rules;
    }

    /**
     * Returns the effective fingering values for a source staff entry. This includes native MusicXML
     * fingerings until they are replaced with setFingeringValues().
     */
    public getFingeringValues(sourceStaffEntry: SourceStaffEntry): string[] {
        if (!sourceStaffEntry) {
            return [];
        }
        const values: string[] = [];
        for (const voiceEntry of sourceStaffEntry.VoiceEntries) {
            for (const instruction of voiceEntry.TechnicalInstructions) {
                if (instruction.type === TechnicalInstructionType.Fingering && instruction.value) {
                    values.push(instruction.value);
                }
            }
        }
        return values;
    }

    /** Returns the original MusicXML fingering values even while an in-memory override is active. */
    public getNativeFingeringValues(sourceStaffEntry: SourceStaffEntry): string[] {
        if (!sourceStaffEntry) {
            return [];
        }
        const instructions: TechnicalInstruction[] = this.fingeringOverrides.has(sourceStaffEntry) ?
            this.nativeFingeringBackups.get(sourceStaffEntry) ?? [] : this.getFingeringInstructions(sourceStaffEntry);
        return instructions.map((instruction: TechnicalInstruction) => instruction.value).filter((value: string) => !!value);
    }

    /**
     * Replaces all fingerings at one staff entry and redraws only its fingering labels.
     *
     * Passing an array (including an empty array) creates an in-memory override. Passing undefined
     * removes the override and restores the original MusicXML fingerings. No layout or full-sheet
     * render is performed, so note SVG nodes, event handlers, and scroll position remain untouched.
     * Live visual updates are supported by the SVG backend; Canvas reflects the source change on its
     * next render.
     */
    public setFingeringValues(sourceStaffEntry: SourceStaffEntry, values: string[] = undefined): boolean {
        if (!sourceStaffEntry) {
            return false;
        }

        if (values === undefined) {
            if (!this.fingeringOverrides.has(sourceStaffEntry)) {
                return true;
            }
            this.removeFingeringInstructions(sourceStaffEntry);
            for (const instruction of this.nativeFingeringBackups.get(sourceStaffEntry) ?? []) {
                const voiceEntry: VoiceEntry = instruction.sourceNote?.ParentVoiceEntry ?? sourceStaffEntry.VoiceEntries[0];
                if (voiceEntry && voiceEntry.TechnicalInstructions.indexOf(instruction) < 0) {
                    voiceEntry.TechnicalInstructions.push(instruction);
                }
                if (instruction.sourceNote) {
                    instruction.sourceNote.Fingering = instruction;
                }
            }
            this.fingeringOverrides.delete(sourceStaffEntry);
            this.nativeFingeringBackups.delete(sourceStaffEntry);
        } else {
            const normalizedValues: string[] = values
                .filter((value: string) => /^[1-5]$/.test(value))
                .filter((value: string, index: number, all: string[]) => all.indexOf(value) === index);
            const currentValues: string[] = this.getFingeringValues(sourceStaffEntry);
            if (this.fingeringOverrides.has(sourceStaffEntry) &&
                currentValues.length === normalizedValues.length &&
                currentValues.every((value: string, index: number) => value === normalizedValues[index])) {
                return true;
            }

            const sourceNote: Note = sourceStaffEntry.VoiceEntries
                .filter(voiceEntry => !voiceEntry.IsGrace)
                .reduce((notes, voiceEntry) => notes.concat(voiceEntry.Notes), [])
                .find(note => !!note.Pitch);
            if (!sourceNote && normalizedValues.length > 0) {
                return false;
            }

            if (!this.fingeringOverrides.has(sourceStaffEntry)) {
                this.nativeFingeringBackups.set(sourceStaffEntry, this.getFingeringInstructions(sourceStaffEntry));
            }
            this.removeFingeringInstructions(sourceStaffEntry);
            for (const value of normalizedValues) {
                const instruction: TechnicalInstruction = new TechnicalInstruction();
                instruction.type = TechnicalInstructionType.Fingering;
                instruction.value = value;
                instruction.sourceNote = sourceNote;
                sourceNote.ParentVoiceEntry.TechnicalInstructions.push(instruction);
                sourceNote.Fingering = instruction;
            }
            this.fingeringOverrides.add(sourceStaffEntry);
        }

        const graphicalStaffEntry: GraphicalStaffEntry = this.graphic?.GetGraphicalFromSourceStaffEntry(sourceStaffEntry);
        if (graphicalStaffEntry && this.backendType === BackendType.SVG) {
            this.redrawFingeringLabels(graphicalStaffEntry, this.getFingeringInstructions(sourceStaffEntry));
        }
        return true;
    }

    private getFingeringInstructions(sourceStaffEntry: SourceStaffEntry): TechnicalInstruction[] {
        const result: TechnicalInstruction[] = [];
        for (const voiceEntry of sourceStaffEntry.VoiceEntries) {
            for (const instruction of voiceEntry.TechnicalInstructions) {
                if (instruction.type === TechnicalInstructionType.Fingering) {
                    result.push(instruction);
                }
            }
        }
        return result;
    }

    private removeFingeringInstructions(sourceStaffEntry: SourceStaffEntry): void {
        for (const voiceEntry of sourceStaffEntry.VoiceEntries) {
            for (let index: number = voiceEntry.TechnicalInstructions.length - 1; index >= 0; index--) {
                const instruction: TechnicalInstruction = voiceEntry.TechnicalInstructions[index];
                if (instruction.type !== TechnicalInstructionType.Fingering) {
                    continue;
                }
                voiceEntry.TechnicalInstructions.splice(index, 1);
                if (instruction.sourceNote?.Fingering === instruction) {
                    instruction.sourceNote.Fingering = undefined;
                }
            }
        }
    }

    private redrawFingeringLabels(graphicalStaffEntry: GraphicalStaffEntry,
                                  instructions: TechnicalInstruction[]): void {
        const oldLabels: GraphicalLabel[] = graphicalStaffEntry.FingeringEntries ?? [];
        const sourceStaffEntry: SourceStaffEntry = graphicalStaffEntry.sourceStaffEntry;
        const renderedAnchorY: number = oldLabels[0] && !this.inPlaceFingeringLabels.has(oldLabels[0]) ?
            oldLabels[0].PositionAndShape.RelativePosition.y : undefined;
        if (renderedAnchorY !== undefined) {
            this.fingeringAnchors.set(sourceStaffEntry, renderedAnchorY);
        }
        const anchorY: number = renderedAnchorY ?? this.fingeringAnchors.get(sourceStaffEntry);
        for (const oldLabel of oldLabels) {
            if (oldLabel.SVGNode?.parentNode) {
                oldLabel.SVGNode.parentNode.removeChild(oldLabel.SVGNode);
            }
            const parent: BoundingBox = oldLabel.PositionAndShape.Parent;
            const childIndex: number = parent?.ChildElements.indexOf(oldLabel.PositionAndShape) ?? -1;
            if (childIndex >= 0) {
                parent.ChildElements.splice(childIndex, 1);
            }
        }
        graphicalStaffEntry.FingeringEntries = [];
        if (instructions.length === 0 || !this.drawer) {
            return;
        }

        const measure: GraphicalMeasure = graphicalStaffEntry.parentMeasure;
        const systemGroup: SVGGElement = this.findSystemGroup(measure.ParentStaffLine?.ParentMusicSystem);
        let placement: PlacementEnum = this.rules.FingeringPosition;
        if (placement === PlacementEnum.NotYetDefined || placement === PlacementEnum.AboveOrBelow) {
            placement = measure.isUpperStaffOfInstrument() ? PlacementEnum.Above : PlacementEnum.Below;
        }
        if (placement === PlacementEnum.Left || placement === PlacementEnum.Right) {
            return;
        }

        const staffLine: StaffLine = measure.ParentStaffLine;
        const staffEntryX: number = graphicalStaffEntry.PositionAndShape.RelativePosition.x +
            measure.PositionAndShape.RelativePosition.x;
        const orderedInstructions: TechnicalInstruction[] = instructions.slice();
        if (placement === PlacementEnum.Below) {
            orderedInstructions.reverse();
        }

        let previousBoundary: number = undefined;
        for (let index: number = 0; index < orderedInstructions.length; index++) {
            const instruction: TechnicalInstruction = orderedInstructions[index];
            const alignment: TextAlignmentEnum = placement === PlacementEnum.Above ?
                TextAlignmentEnum.CenterBottom : TextAlignmentEnum.CenterTop;
            const label: Label = new Label(instruction.value, alignment);
            if (instruction.fontFamily) {
                label.fontFamily = instruction.fontFamily;
            }
            const graphicalLabel: GraphicalLabel = new GraphicalLabel(
                label, this.rules.FingeringTextSize, alignment, this.rules, staffLine.PositionAndShape);
            graphicalLabel.PositionAndShape.RelativePosition.x = staffEntryX;
            graphicalLabel.setLabelPositionAndShapeBorders();

            if (index === 0 && anchorY !== undefined) {
                graphicalLabel.PositionAndShape.RelativePosition.y = anchorY;
            } else if (index > 0) {
                graphicalLabel.PositionAndShape.RelativePosition.y = previousBoundary +
                    (placement === PlacementEnum.Above ? -this.rules.FingeringPaddingY : this.rules.FingeringPaddingY);
            } else {
                const labelLeft: number = staffEntryX + graphicalLabel.PositionAndShape.BorderMarginLeft;
                const labelRight: number = staffEntryX + graphicalLabel.PositionAndShape.BorderMarginRight;
                const entryLeft: number = staffEntryX + graphicalStaffEntry.PositionAndShape.BorderLeft;
                const entryRight: number = staffEntryX + graphicalStaffEntry.PositionAndShape.BorderRight;
                const marginLeft: number = Math.min(labelLeft, entryLeft);
                const marginRight: number = Math.max(labelRight, entryRight);
                const skyline: SkyBottomLineCalculator = staffLine.SkyBottomLineCalculator;
                const furthest: number = placement === PlacementEnum.Above ?
                    skyline.getSkyLineMinInRange(marginLeft, marginRight) :
                    skyline.getBottomLineMaxInRange(marginLeft, marginRight);
                const offset: number = this.rules.FingeringOffsetY + (placement === PlacementEnum.Above ? 0.1 : 0);
                graphicalLabel.PositionAndShape.RelativePosition.y = furthest +
                    (placement === PlacementEnum.Above ? -offset : offset);
                graphicalLabel.PositionAndShape.RelativePosition.y = this.getFingeringYNextToNotes(
                    staffLine, placement, graphicalLabel, marginLeft, marginRight, offset, orderedInstructions.length) ??
                    graphicalLabel.PositionAndShape.RelativePosition.y;
            }
            if (anchorY === undefined) {
                this.inPlaceFingeringLabels.add(graphicalLabel);
            }

            graphicalLabel.PositionAndShape.calculateBoundingBox();
            graphicalLabel.PositionAndShape.setAbsolutePositionFromParent();
            previousBoundary = graphicalLabel.PositionAndShape.RelativePosition.y +
                (placement === PlacementEnum.Above ? graphicalLabel.PositionAndShape.BorderTop :
                    graphicalLabel.PositionAndShape.BorderBottom);
            if (systemGroup) {
                const node: Node = this.drawer.drawLabel(graphicalLabel, GraphicalLayers.Notes);
                if (node && node.parentNode !== systemGroup) {
                    systemGroup.appendChild(node);
                }
                graphicalLabel.SVGNode = node;
            }
            graphicalStaffEntry.FingeringEntries.push(graphicalLabel);
        }
    }

    private findSystemGroup(system: MusicSystem): SVGGElement {
        const page: GraphicalMusicPage = system?.Parent;
        const systemIndex: number = page?.MusicSystems.indexOf(system) ?? -1;
        if (systemIndex < 0) {
            return undefined;
        }
        const key: string = `${page.PageNumber}:${systemIndex}`;
        const groups: SVGGElement[] = this.drawer?.SystemGroups ?? [];
        for (let index: number = groups.length - 1; index >= 0; index--) {
            if (groups[index].getAttribute("data-osmd-system-key") === key) {
                return groups[index];
            }
        }
        return undefined;
    }

    private getFingeringYNextToNotes(staffLine: StaffLine, placement: PlacementEnum, label: GraphicalLabel,
                                     left: number, right: number, offset: number, count: number): number {
        const skyline: SkyBottomLineCalculator = staffLine.SkyBottomLineCalculator;
        const above: boolean = placement === PlacementEnum.Above;
        const fingeringLine: number[] = above ? skyline?.FingeringSkyLine : skyline?.FingeringBottomLine;
        const slurLine: number[] = above ? skyline?.SlurSkyLine : skyline?.SlurBottomLine;
        if (!fingeringLine || !slurLine) {
            return undefined;
        }
        const current: number = above ? skyline.getSkyLineMinInRange(left, right) : skyline.getBottomLineMaxInRange(left, right);
        const afterSlurs: number = above ? skyline.getMinInLineRange(slurLine, left, right) :
            skyline.getMaxInLineRange(slurLine, left, right);
        if (current !== afterSlurs) {
            return undefined;
        }
        const y: number = above ? skyline.getMinInLineRange(fingeringLine, left, right) - offset :
            skyline.getMaxInLineRange(fingeringLine, left, right) + offset;
        const currentY: number = current + (above ? -offset : offset);
        if (!Number.isFinite(y) || (above ? y <= currentY : y >= currentY)) {
            return undefined;
        }

        const box: BoundingBox = label.PositionAndShape;
        const padding: number = 0.2;
        const stackHeight: number = count * (box.BorderBottom - box.BorderTop) + (count - 1) * this.rules.FingeringPaddingY;
        const top: number = (above ? y + box.BorderBottom - stackHeight : y + box.BorderTop) - padding;
        const bottom: number = top + stackHeight + 2 * padding;
        const minX: number = box.RelativePosition.x + box.BorderLeft - padding;
        const maxX: number = box.RelativePosition.x + box.BorderRight + padding;
        for (const line of staffLine.ParentMusicSystem.StaffLines) {
            if (line.GraphicalGlissandi.length > 0) {
                return undefined;
            }
            const dx: number = line.PositionAndShape.RelativePosition.x - staffLine.PositionAndShape.RelativePosition.x;
            const dy: number = line.PositionAndShape.RelativePosition.y - staffLine.PositionAndShape.RelativePosition.y;
            for (const slur of line.GraphicalSlurs) {
                const points: PointF2D[] = [slur.bezierStartPt, slur.bezierStartControlPt, slur.bezierEndControlPt, slur.bezierEndPt];
                if (points.some((point: PointF2D) => !point || !Number.isFinite(point.x) || !Number.isFinite(point.y))) {
                    continue;
                }
                for (let step: number = 0; step <= 40; step++) {
                    const t: number = step / 40;
                    const s: number = 1 - t;
                    const weights: number[] = [s * s * s, 3 * s * s * t, 3 * s * t * t, t * t * t];
                    const x: number = weights.reduce((sum: number, w: number, i: number) => sum + w * points[i].x, 0) + dx;
                    if (x < minX || x > maxX) {
                        continue;
                    }
                    const curveY: number = weights.reduce((sum: number, w: number, i: number) => sum + w * points[i].y, 0) + dy;
                    const thickness: number = 0.05 * (weights[0] + weights[3]) + 0.3 * (weights[1] + weights[2]);
                    const outerY: number = curveY + (slur.placement === PlacementEnum.Above ? -thickness : thickness);
                    if (Math.max(curveY, outerY) >= top && Math.min(curveY, outerY) <= bottom) {
                        return undefined;
                    }
                }
            }
        }
        return y;
    }

    /** Returns the version of OSMD this object is built from (the version you are using). */
    public get Version(): string {
        return this.version;
    }
    //#endregion

    /**
     * Lowers the opacity of every rendered instance of the staff that matches the provided global staff index.
     * A staff can appear on multiple systems/pages; all of them are updated in-place via the SVG DOM.
     * @param staffIndex Global staff index (`Staff.idInMusicSheet`)
     * @param opacity Target opacity in the range [0, 1]
     */
    public blurStaff(staffIndex: number, opacity: number = 0.3): void {
        this.setStaffOpacity(staffIndex, opacity, 0.3);
    }

    /**
     * Restores the opacity of a staff (across all systems/pages) back to 1.0.
     * @param staffIndex Global staff index (`Staff.idInMusicSheet`)
     */
    public restoreStaff(staffIndex: number): void {
        this.setStaffOpacity(staffIndex, 1.0, 1.0);
    }

    public blurVoice(voiceId: number, opacity: number = 0.2): void {
        if (!this.sheet || !this.graphic) {
            return;
        }

        for (const instrument of this.sheet.Instruments) {
            for (const staff of instrument.Staves) {
                for (const voice of staff.Voices) {
                    if (voice.VoiceId === voiceId) {
                        for (const voiceEntry of voice.VoiceEntries) {
                            for (const note of voiceEntry.Notes) {
                                const gNote: GraphicalNote = this.rules.GNote(note);
                                if (gNote) {
                                    gNote.setOpacity(opacity);
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    public blurVoices(voiceIds: number[], opacity: number = 0.2): void {
        if (!this.sheet || !this.graphic) {
            return;
        }

        for (const instrument of this.sheet.Instruments) {
            for (const staff of instrument.Staves) {
                for (const voice of staff.Voices) {
                    if (voiceIds?.includes(voice.VoiceId)) {
                        for (const voiceEntry of voice.VoiceEntries) {
                            for (const note of voiceEntry.Notes) {
                                const gNote: GraphicalNote = this.rules.GNote(note);
                                if (gNote) {
                                    gNote.setOpacity(opacity);
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    public blurAllVoicesExceptVoices(voiceIds: number[], opacity: number = 0.2): void {
        if (!this.sheet || !this.graphic) {
            return;
        }

        for (const instrument of this.sheet.Instruments) {
            for (const staff of instrument.Staves) {
                for (const voice of staff.Voices) {
                    if (!voiceIds?.includes(voice.VoiceId)) {
                        for (const voiceEntry of voice.VoiceEntries) {
                            for (const note of voiceEntry.Notes) {
                                const gNote: GraphicalNote = this.rules.GNote(note);
                                if (gNote) {
                                    gNote.setOpacity(opacity);
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    public resetOpacity(): void {
        if (!this.sheet || !this.graphic) {
            return;
        }

        for (const instrument of this.sheet.Instruments) {
            for (const staff of instrument.Staves) {
                for (const voice of staff.Voices) {
                    for (const voiceEntry of voice.VoiceEntries) {
                        for (const note of voiceEntry.Notes) {
                            const gNote: GraphicalNote = this.rules.GNote(note);
                            if (gNote) {
                                gNote.setOpacity(1.0);
                            }
                        }
                    }
                }
            }
        }
    }

    /**
     * Read-ahead helper: set the opacity of notes within a single measure. This is used by the
     * sight-reading "read ahead" mode to hide a rolling window of beats so the player has to read
     * ahead of where they are currently playing.
     * @param measureListIndex 0-based index into the measure list (`SourceMeasure.measureListIndex`).
     * @param hiddenBeats 0 hides the whole measure (legacy behavior); N > 0 hides a window of N
     * quarter-note beats starting at `fromBeatInMeasure`. Notes outside the window are made fully
     * visible so the window can slide forward across repeated calls.
     * @param opacity opacity applied to the hidden notes (default 0 = fully hidden).
     * @param fromBeatInMeasure quarter-note beats from the measure start where the hidden window
     * begins (default 0). Ignored when `hiddenBeats <= 0`.
     */
    public setReadAheadMeasureOpacity(
        measureListIndex: number,
        hiddenBeats: number = 0,
        opacity: number = 0,
        fromBeatInMeasure: number = 0
    ): void {
        if (!this.sheet || !this.graphic) {
            return;
        }
        // relInMeasureTimestamp.RealValue is expressed in whole notes (a quarter note = 0.25),
        // so beats map to whole-note units via * 0.25.
        const quarterNoteFraction: number = 0.25;
        const windowStart: number = hiddenBeats > 0 ? fromBeatInMeasure * quarterNoteFraction : 0;
        const windowEnd: number = hiddenBeats > 0 ? (fromBeatInMeasure + hiddenBeats) * quarterNoteFraction : Number.POSITIVE_INFINITY;
        const epsilon: number = 1e-6;

        for (const staffEntry of this.getReadAheadStaffEntriesForMeasure(measureListIndex)) {
            const inMeasureTime: number = staffEntry.relInMeasureTimestamp?.RealValue ?? 0;
            // hiddenBeats <= 0: whole measure hidden. Otherwise hide only [windowStart, windowEnd).
            const inWindow: boolean =
                hiddenBeats <= 0 ||
                (inMeasureTime >= windowStart - epsilon && inMeasureTime < windowEnd - epsilon);
            const targetOpacity: number = inWindow ? opacity : 1.0;
            for (const graphicalVoiceEntry of staffEntry.graphicalVoiceEntries ?? []) {
                for (const graphicalNote of graphicalVoiceEntry?.notes ?? []) {
                    if (graphicalNote) {
                        graphicalNote.setOpacity(targetOpacity);
                        if (inWindow) {
                            this.readAheadOpacityTouchedGraphicalNotes.add(graphicalNote);
                        } else {
                            this.readAheadOpacityTouchedGraphicalNotes.delete(graphicalNote);
                        }
                    }
                }
            }
            for (const fingeringEntry of staffEntry.FingeringEntries ?? []) {
                this.setReadAheadGraphicalLabelOpacity(fingeringEntry, targetOpacity);
            }
        }
    }

    /**
     * Read-ahead helper: hide a beat range within a single measure without restoring notes outside
     * that range. This supports cumulative read-ahead where each playback beat only hides newly
     * reached future notes.
     * @param measureListIndex 0-based index into the measure list (`SourceMeasure.measureListIndex`).
     * @param fromBeatInMeasure quarter-note beats from the measure start where hiding begins.
     * @param beatCount number of quarter-note beats to hide from `fromBeatInMeasure`.
     * @param opacity opacity applied to hidden notes (default 0 = fully hidden).
     */
    public hideReadAheadMeasureBeatRange(
        measureListIndex: number,
        fromBeatInMeasure: number,
        beatCount: number,
        opacity: number = 0
    ): void {
        if (!this.sheet || !this.graphic || beatCount <= 0) {
            return;
        }
        const quarterNoteFraction: number = 0.25;
        const windowStart: number = Math.max(0, fromBeatInMeasure) * quarterNoteFraction;
        const windowEnd: number = Math.max(0, fromBeatInMeasure + beatCount) * quarterNoteFraction;
        const epsilon: number = 1e-6;

        const staffEntries: GraphicalStaffEntry[] = this.getReadAheadStaffEntriesForMeasure(measureListIndex);
        const sorted: boolean = !this.readAheadUnsortedMeasures.has(measureListIndex);
        const firstIndex: number = sorted ? this.firstReadAheadStaffEntryAtOrAfter(staffEntries, windowStart - epsilon) : 0;
        for (let index: number = firstIndex; index < staffEntries.length; index++) {
            const staffEntry: GraphicalStaffEntry = staffEntries[index];
            const inMeasureTime: number = staffEntry.relInMeasureTimestamp?.RealValue ?? 0;
            if (sorted && inMeasureTime >= windowEnd - epsilon) {
                break;
            }
            if (inMeasureTime < windowStart - epsilon || inMeasureTime >= windowEnd - epsilon) {
                continue;
            }
            for (const graphicalVoiceEntry of staffEntry.graphicalVoiceEntries ?? []) {
                for (const graphicalNote of graphicalVoiceEntry?.notes ?? []) {
                    if (graphicalNote) {
                        graphicalNote.setOpacity(opacity);
                        this.readAheadOpacityTouchedGraphicalNotes.add(graphicalNote);
                    }
                }
            }
            for (const fingeringEntry of staffEntry.FingeringEntries ?? []) {
                this.setReadAheadGraphicalLabelOpacity(fingeringEntry, opacity);
            }
        }
    }

    /** Restores the opacity of every note hidden by {@link setReadAheadMeasureOpacity}. */
    public resetReadAheadOpacity(): void {
        if (this.readAheadOpacityTouchedGraphicalNotes.size === 0 && this.readAheadOpacityTouchedLabels.size === 0) {
            return;
        }
        for (const graphicalNote of this.readAheadOpacityTouchedGraphicalNotes) {
            graphicalNote.setOpacity(1.0);
        }
        for (const label of this.readAheadOpacityTouchedLabels) {
            label.readAheadOpacity = 1;
            if (label.SVGNode instanceof Element && label.SVGNode.isConnected) {
                label.SVGNode.setAttribute("opacity", "1");
            }
        }
        this.readAheadOpacityTouchedGraphicalNotes.clear();
        this.readAheadOpacityTouchedLabels.clear();
    }

    private invalidateReadAheadStaffEntryIndex(): void {
        this.readAheadStaffEntryIndexGeneration = 0;
        this.readAheadStaffEntriesByMeasure.clear();
        this.readAheadUnsortedMeasures.clear();
    }

    private getReadAheadStaffEntriesForMeasure(measureListIndex: number): GraphicalStaffEntry[] {
        const containers: VerticalGraphicalStaffEntryContainer[] = this.graphic?.VerticalGraphicalStaffEntryContainers;
        if (!containers) {
            return [];
        }
        if (this.readAheadStaffEntryIndexGeneration !== this.graphic.LayoutGeneration) {
            this.invalidateReadAheadStaffEntryIndex();
            for (const container of containers) {
                for (const staffEntry of container?.StaffEntries ?? []) {
                    const index: number = staffEntry?.parentMeasure?.parentSourceMeasure?.measureListIndex;
                    if (index === undefined) {
                        continue;
                    }
                    const entries: GraphicalStaffEntry[] = this.readAheadStaffEntriesByMeasure.get(index) ?? [];
                    const previous: GraphicalStaffEntry = entries[entries.length - 1];
                    if (previous &&
                        !((previous.relInMeasureTimestamp?.RealValue ?? 0) <= (staffEntry.relInMeasureTimestamp?.RealValue ?? 0))) {
                        this.readAheadUnsortedMeasures.add(index);
                    }
                    entries.push(staffEntry);
                    this.readAheadStaffEntriesByMeasure.set(index, entries);
                }
            }
            this.readAheadStaffEntryIndexGeneration = this.graphic.LayoutGeneration;
        }
        return this.readAheadStaffEntriesByMeasure.get(measureListIndex) ?? [];
    }

    private firstReadAheadStaffEntryAtOrAfter(staffEntries: GraphicalStaffEntry[], inMeasureTime: number): number {
        let low: number = 0;
        let high: number = staffEntries.length;
        while (low < high) {
            const mid: number = Math.floor((low + high) / 2);
            if ((staffEntries[mid].relInMeasureTimestamp?.RealValue ?? 0) < inMeasureTime) {
                low = mid + 1;
            } else {
                high = mid;
            }
        }
        return low;
    }

    private setReadAheadGraphicalLabelOpacity(label: GraphicalLabel, opacity: number): void {
        if (!label) {
            return;
        }
        label.readAheadOpacity = opacity;
        if (label.SVGNode instanceof Element) {
            label.SVGNode.setAttribute("opacity", opacity.toString());
        }
        if (opacity < 1.0) {
            this.readAheadOpacityTouchedLabels.add(label);
        } else {
            this.readAheadOpacityTouchedLabels.delete(label);
        }
    }

    private syncInteractiveRangeSelection(): void {
        if (!this.rangeSelection.enabled || !this.graphic || !this.drawer?.Backends?.length) {
            this.detachRangeSelectionListeners();
            this.removeRangeSelectionOverlay();
            return;
        }

        this.ensureRangeSelectionOverlay();
        this.updateRangeSelectionOverlayStyles();
        // hideSelectionRange means display-only (playback, preset segment): keep gray-out/mask but block drag/tap.
        if (this.shouldHideSelectionRangeVisuals()) {
            this.detachRangePointerListeners();
        } else {
            this.attachRangeSelectionListeners();
        }
        this.attachRangeViewportListeners();
        this.renderRangeSelection();
    }

    private attachRangeSelectionListeners(): void {
        const currentElements: HTMLElement[] = this.drawer.Backends
            .map((backend: VexFlowBackend) => backend.getRenderElement())
            .filter((element: HTMLElement) => !!element);
        const sameBinding: boolean = currentElements.length === this.rangeInteractionBoundElements.length
            && currentElements.every((element: HTMLElement, index: number) => this.rangeInteractionBoundElements[index] === element);
        if (sameBinding) {
            return;
        }
        this.detachRangeSelectionListeners();
        for (const element of currentElements) {
            element.addEventListener("pointermove", this.rangePointerMoveListener, { passive: false });
            element.addEventListener("pointerdown", this.rangePointerDownListener);
            element.addEventListener("pointerleave", this.rangePointerLeaveListener, { passive: true });
            this.rangeInteractionBoundElements.push(element);
        }
        if (this.rangeInteractionBoundElements.length > 0) {
            window.addEventListener("pointerup", this.rangePointerUpListener);
            window.addEventListener("pointercancel", this.rangePointerCancelListener);
        }
    }

    private detachRangePointerListeners(): void {
        this.resetTouchGestureState();
        this.pendingTouchRangeStartAnchor = undefined;
        this.emitRangeHandleDragging(false);
        this.releaseRangeDragPointerCapture();
        if (this.rangePointerMoveAnimationFrameId !== 0) {
            window.cancelAnimationFrame(this.rangePointerMoveAnimationFrameId);
            this.rangePointerMoveAnimationFrameId = 0;
        }
        this.pendingRangePointerMoveAnchor = undefined;
        for (const element of this.rangeInteractionBoundElements) {
            element.removeEventListener("pointermove", this.rangePointerMoveListener);
            element.removeEventListener("pointerdown", this.rangePointerDownListener);
            element.removeEventListener("pointerleave", this.rangePointerLeaveListener);
            element.style.removeProperty("cursor");
        }
        this.rangeInteractionBoundElements = [];
        window.removeEventListener("pointerup", this.rangePointerUpListener);
        window.removeEventListener("pointercancel", this.rangePointerCancelListener);
    }

    private detachRangeSelectionListeners(): void {
        this.cancelPendingRangeOpacityUpdate();
        this.detachRangeViewportListeners();
        this.detachRangePointerListeners();
        if (this.rangeViewportUpdateAnimationFrameId !== 0) {
            window.cancelAnimationFrame(this.rangeViewportUpdateAnimationFrameId);
            this.rangeViewportUpdateAnimationFrameId = 0;
        }
        if (this.rangeViewportSettleUpdateTimeoutId !== 0) {
            window.clearTimeout(this.rangeViewportSettleUpdateTimeoutId);
            this.rangeViewportSettleUpdateTimeoutId = 0;
        }
    }

    private attachRangeViewportListeners(): void {
        const scrollTarget: HTMLElement | Window = this.getRangeSelectionScrollContainer();
        if (this.rangeViewportScrollTarget === scrollTarget) {
            return;
        }
        this.detachRangeViewportListeners();
        this.rangeViewportScrollTarget = scrollTarget;
        this.rangeViewportScrollTarget.addEventListener("scroll", this.rangeViewportUpdateListener, { passive: true });
        // Capture window-level scroll as a fallback when the app scrolls in wrappers above the detected target.
        window.addEventListener("scroll", this.rangeViewportUpdateListener, { passive: true, capture: true });
        window.addEventListener("resize", this.rangeViewportUpdateListener, { passive: true });
    }

    private detachRangeViewportListeners(): void {
        if (this.rangeViewportScrollTarget) {
            this.rangeViewportScrollTarget.removeEventListener("scroll", this.rangeViewportUpdateListener);
            this.rangeViewportScrollTarget = undefined;
        }
        window.removeEventListener("scroll", this.rangeViewportUpdateListener, true);
        window.removeEventListener("resize", this.rangeViewportUpdateListener);
        if (this.rangeViewportSettleUpdateTimeoutId !== 0) {
            window.clearTimeout(this.rangeViewportSettleUpdateTimeoutId);
            this.rangeViewportSettleUpdateTimeoutId = 0;
        }
    }

    private scheduleRangeViewportUpdate(): void {
        if (this.rangeViewportUpdateAnimationFrameId !== 0) {
            return;
        }
        this.rangeViewportUpdateAnimationFrameId = window.requestAnimationFrame((): void => {
            this.rangeViewportUpdateAnimationFrameId = 0;
            if (!this.rangeSelection.enabled || !this.rangeInteractionOverlay) {
                return;
            }
            if (this.isRangeDragging) {
                // Active drag renders via the pointer-move path; nothing to do for scroll here.
                return;
            }
            if (this.dragStartAnchor && this.dragCurrentAnchor) {
                if (!this.shouldUseMaskGrayOut()) {
                    // Committed range with per-note opacity: extend gray-out to newly visible systems.
                    // Mask-based gray-out is container-relative and needs no scroll refresh.
                    this.updateRangeSelectionViewport(false);
                }
                return;
            }
            if (this.pendingTouchRangeStartAnchor || this.hoverAnchor) {
                this.renderRangeSelection();
            }
        });
    }

    private scheduleRangeViewportSettleUpdate(): void {
        if (this.rangeViewportSettleUpdateTimeoutId !== 0) {
            window.clearTimeout(this.rangeViewportSettleUpdateTimeoutId);
        }
        this.rangeViewportSettleUpdateTimeoutId = window.setTimeout((): void => {
            this.rangeViewportSettleUpdateTimeoutId = 0;
            if (!this.rangeSelection.enabled || !this.rangeInteractionOverlay) {
                return;
            }
            if (this.isRangeDragging || !this.dragStartAnchor || !this.dragCurrentAnchor) {
                return;
            }
            if (this.shouldUseMaskGrayOut()) {
                return;
            }
            // Scrolling settled: apply a full-fidelity gray-out (including decorations) for the
            // final viewport, again without rebuilding the overlay or action buttons.
            this.updateRangeSelectionViewport(true);
        }, 120);
    }

    private ensureRangeSelectionOverlay(): void {
        if (this.rangeInteractionOverlay) {
            // Backend refresh can remove all container children, which detaches this node.
            // If the reference exists but the node is no longer connected, re-attach it.
            const overlayDetached: boolean = !this.rangeInteractionOverlay.isConnected
                || this.rangeInteractionOverlay.parentElement !== this.container;
            if (!overlayDetached) {
                return;
            }
            this.rangeInteractionOverlay.style.position = "absolute";
            this.rangeInteractionOverlay.style.left = "0";
            this.rangeInteractionOverlay.style.top = "0";
            this.rangeInteractionOverlay.style.right = "0";
            this.rangeInteractionOverlay.style.bottom = "0";
            this.rangeInteractionOverlay.style.pointerEvents = "none";
            this.rangeInteractionOverlay.style.zIndex = this.getSelectionOverlayZIndex().toString();
            this.container.appendChild(this.rangeInteractionOverlay);
            return;
        }
        if (window.getComputedStyle(this.container).position === "static") {
            this.container.style.position = "relative";
        }
        this.rangeInteractionOverlay = document.createElement("div");
        this.rangeInteractionOverlay.className = "osmd-range-selection-overlay";
        this.rangeInteractionOverlay.style.position = "absolute";
        this.rangeInteractionOverlay.style.left = "0";
        this.rangeInteractionOverlay.style.top = "0";
        this.rangeInteractionOverlay.style.right = "0";
        this.rangeInteractionOverlay.style.bottom = "0";
        this.rangeInteractionOverlay.style.pointerEvents = "none";
        this.rangeInteractionOverlay.style.zIndex = this.getSelectionOverlayZIndex().toString();
        this.container.appendChild(this.rangeInteractionOverlay);
        this.outsideMaskLayer = undefined;
        this.rangeChromeLayer = undefined;
        this.outsideMaskPool.length = 0;
        this.dragHandleLines[0] = undefined;
        this.dragHandleLines[1] = undefined;
    }

    private ensureOutsideMaskLayer(): HTMLDivElement {
        this.ensureRangeSelectionOverlay();
        if (!this.outsideMaskLayer || !this.outsideMaskLayer.isConnected) {
            this.outsideMaskLayer = document.createElement("div");
            this.outsideMaskLayer.className = "osmd-range-outside-mask-layer";
            this.outsideMaskLayer.style.position = "absolute";
            this.outsideMaskLayer.style.left = "0";
            this.outsideMaskLayer.style.top = "0";
            this.outsideMaskLayer.style.right = "0";
            this.outsideMaskLayer.style.bottom = "0";
            this.outsideMaskLayer.style.pointerEvents = "none";
            if (this.rangeChromeLayer?.isConnected) {
                this.rangeInteractionOverlay.insertBefore(this.outsideMaskLayer, this.rangeChromeLayer);
            } else {
                this.rangeInteractionOverlay.appendChild(this.outsideMaskLayer);
            }
        }
        return this.outsideMaskLayer;
    }

    private ensureRangeChromeLayer(): HTMLDivElement {
        this.ensureRangeSelectionOverlay();
        if (!this.rangeChromeLayer || !this.rangeChromeLayer.isConnected) {
            this.rangeChromeLayer = document.createElement("div");
            this.rangeChromeLayer.className = "osmd-range-chrome-layer";
            this.rangeChromeLayer.style.position = "absolute";
            this.rangeChromeLayer.style.left = "0";
            this.rangeChromeLayer.style.top = "0";
            this.rangeChromeLayer.style.right = "0";
            this.rangeChromeLayer.style.bottom = "0";
            this.rangeChromeLayer.style.pointerEvents = "none";
            this.rangeInteractionOverlay.appendChild(this.rangeChromeLayer);
            this.dragHandleLines[0] = undefined;
            this.dragHandleLines[1] = undefined;
        }
        return this.rangeChromeLayer;
    }

    private clearRangeChromeLayer(): void {
        if (!this.rangeChromeLayer) {
            return;
        }
        this.rangeChromeLayer.innerHTML = "";
        this.dragHandleLines[0] = undefined;
        this.dragHandleLines[1] = undefined;
    }

    private updateRangeSelectionOverlayStyles(): void {
        if (!this.rangeInteractionOverlay) {
            return;
        }
        this.rangeInteractionOverlay.style.zIndex = this.getSelectionOverlayZIndex().toString();
    }

    private removeRangeSelectionOverlay(): void {
        if (!this.rangeInteractionOverlay) {
            return;
        }
        this.rangeInteractionOverlay.remove();
        this.rangeInteractionOverlay = undefined;
        this.outsideMaskLayer = undefined;
        this.rangeChromeLayer = undefined;
        this.outsideMaskPool.length = 0;
        this.dragHandleLines[0] = undefined;
        this.dragHandleLines[1] = undefined;
        this.maskDragChromePrepared = false;
    }

    private onRangePointerMove(event: PointerEvent): void {
        if (!this.rangeSelection.enabled || this.shouldHideSelectionRangeVisuals()) {
            return;
        }
        if (this.isTouchPointerEvent(event)) {
            this.updateTouchMoveState(event);
            if (event.pointerId !== this.activeTouchPointerId) {
                return;
            }
            if (this.isRangeDragging) {
                this.activeTouchDragClientX = event.clientX;
                this.activeTouchDragClientY = event.clientY;
                event.preventDefault();
            }
            // Let touch gestures default to native page/score scrolling unless we are actively dragging a handle.
            if (!this.isRangeDragging) {
                return;
            }
        }
        const anchor: RangeSelectionAnchor = this.getAnchorFromPointerEvent(event);
        this.updateDesktopRangeCursor(event, anchor);
        if (!anchor) {
            return;
        }
        this.pendingRangePointerMoveAnchor = anchor;
        if (this.rangePointerMoveAnimationFrameId !== 0) {
            return;
        }
        this.rangePointerMoveAnimationFrameId = window.requestAnimationFrame((): void => {
            this.rangePointerMoveAnimationFrameId = 0;
            this.flushRangePointerMove();
        });
    }

    private flushRangePointerMove(): void {
        if (!this.rangeSelection.enabled) {
            this.pendingRangePointerMoveAnchor = undefined;
            return;
        }
        const anchor: RangeSelectionAnchor = this.pendingRangePointerMoveAnchor;
        this.pendingRangePointerMoveAnchor = undefined;
        if (!anchor) {
            return;
        }
        this.hoverAnchor = anchor;
        if (this.isRangeDragging && this.dragStartAnchor) {
            this.dragCurrentAnchor = anchor;
            this.renderRangeSelection();
            this.emitRangeSelection("dragging", this.dragStartAnchor, anchor, true);
            return;
        }
        this.renderRangeSelection();
        this.emitRangeSelection("hover", anchor, anchor, false);
    }

    private onRangePointerDown(event: PointerEvent): void {
        if (!this.rangeSelection.enabled || this.shouldHideSelectionRangeVisuals()) {
            return;
        }
        if (this.isTouchPointerEvent(event)) {
            this.onRangeTouchPointerDown(event);
            return;
        }
        this.pendingTouchRangeStartAnchor = undefined;
        this.captureRangePointer(event);
        const anchor: RangeSelectionAnchor = this.getAnchorFromPointerEvent(event);
        if (!anchor) {
            return;
        }
        const existingSelection: RangeSelectionPayload = this.getRangeSelection();
        const draggedBound: "start" | "end" | undefined = this.getDraggedBoundFromAnchor(anchor, existingSelection);
        if (existingSelection && !this.isAnchorInsideSelection(anchor, existingSelection)) {
            if (draggedBound) {
                this.isRangeDragging = true;
                if (draggedBound === "start") {
                    // Resize start bound (keep end fixed).
                    this.activeDragBound = "start";
                    this.dragStartAnchor = existingSelection.normalizedEnd;
                } else {
                    // Resize end bound (keep start fixed).
                    this.activeDragBound = "end";
                    this.dragStartAnchor = existingSelection.normalizedStart;
                }
                this.dragCurrentAnchor = anchor;
                this.emitRangeHandleDragging(true);
                this.renderRangeSelection();
                this.emitRangeSelection("dragging", this.dragStartAnchor, this.dragCurrentAnchor, true);
                event.preventDefault();
                return;
            }
            this.clearRangeSelection(true);
            event.preventDefault();
            return;
        }
        this.isRangeDragging = true;
        if (existingSelection) {
            if (draggedBound === "start") {
                // Resize start bound (keep end fixed).
                this.activeDragBound = "start";
                this.dragStartAnchor = existingSelection.normalizedEnd;
                this.dragCurrentAnchor = anchor;
            } else if (draggedBound === "end") {
                // Resize end bound (keep start fixed).
                this.activeDragBound = "end";
                this.dragStartAnchor = existingSelection.normalizedStart;
                this.dragCurrentAnchor = anchor;
            } else {
            this.activeDragBound = "both";
            const startDistance: number = Math.abs(anchor.timestampReal - existingSelection.normalizedStart.timestampReal);
            const endDistance: number = Math.abs(anchor.timestampReal - existingSelection.normalizedEnd.timestampReal);
            if (startDistance <= endDistance) {
                // Resize start bound (keep end fixed).
                this.dragStartAnchor = existingSelection.normalizedEnd;
            } else {
                // Resize end bound (keep start fixed).
                this.dragStartAnchor = existingSelection.normalizedStart;
            }
            this.dragCurrentAnchor = anchor;
            }
        } else {
            this.activeDragBound = "both";
            this.dragStartAnchor = anchor;
            this.dragCurrentAnchor = anchor;
        }
        this.emitRangeHandleDragging(this.activeDragBound !== "both");
        this.renderRangeSelection();
        this.emitRangeSelection("dragging", this.dragStartAnchor, this.dragCurrentAnchor, true);
        event.preventDefault();
    }

    private onRangePointerUp(event: PointerEvent): void {
        if (this.isTouchPointerEvent(event)) {
            this.onRangeTouchPointerUp(event);
            return;
        }
        this.releaseRangeDragPointerCapture();
        this.commitRangeDrag(event);
    }

    private onRangePointerCancel(event: PointerEvent): void {
        if (!this.rangeSelection.enabled || !this.isTouchPointerEvent(event)) {
            return;
        }
        if (event.pointerId !== this.activeTouchPointerId) {
            return;
        }
        this.releaseRangeDragPointerCapture();
        this.commitRangeDrag();
        this.resetTouchGestureState();
    }

    private releaseRangeDragPointerCapture(): void {
        if (!this.rangeDragPointerCaptureElement || this.rangeDragPointerId < 0) {
            return;
        }
        const captureElement: Element & {
            hasPointerCapture?: (pointerId: number) => boolean;
            releasePointerCapture?: (pointerId: number) => void;
        } = this.rangeDragPointerCaptureElement as any;
        if (captureElement.hasPointerCapture?.(this.rangeDragPointerId) && captureElement.releasePointerCapture) {
            captureElement.releasePointerCapture(this.rangeDragPointerId);
        }
        this.rangeDragPointerCaptureElement = undefined;
        this.rangeDragPointerId = -1;
    }

    private onRangePointerLeave(event: PointerEvent): void {
        this.clearDesktopRangeCursor(event);
        if (this.isRangeDragging) {
            return;
        }
        const relatedTarget: Node = event.relatedTarget as Node;
        if (relatedTarget && this.rangeInteractionOverlay?.contains(relatedTarget)) {
            return;
        }
        this.hoverAnchor = undefined;
        this.renderRangeSelection();
    }

    private onRangeTouchPointerDown(event: PointerEvent): void {
        if (this.activeTouchPointerId !== -1 && this.activeTouchPointerId !== event.pointerId) {
            return;
        }
        const anchor: RangeSelectionAnchor = this.getAnchorFromPointerEvent(event);
        if (!anchor) {
            return;
        }
        this.activeTouchPointerId = event.pointerId;
        this.activeTouchStartClientX = event.clientX;
        this.activeTouchStartClientY = event.clientY;
        this.activeTouchMoved = false;
        this.activeTouchDownAnchor = anchor;
        this.activeTouchDragClientX = event.clientX;
        this.activeTouchDragClientY = event.clientY;
        this.touchPendingAction = "none";

        const existingSelection: RangeSelectionPayload = this.getRangeSelection();
        const draggedBound: "start" | "end" | undefined = this.getDraggedBoundFromAnchor(anchor, existingSelection, true);
        if (existingSelection) {
            if (draggedBound) {
                this.pendingTouchRangeStartAnchor = undefined;
                this.captureRangePointer(event);
                this.startRangeHandleDrag(existingSelection, draggedBound, anchor);
                event.preventDefault();
                return;
            }
            if (this.isAnchorInsideSelection(anchor, existingSelection)) {
                return;
            }
            this.touchPendingAction = "clearSelection";
            return;
        }
        this.touchPendingAction = "setOrCommit";
    }

    private onRangeTouchPointerUp(event: PointerEvent): void {
        if (this.activeTouchPointerId !== event.pointerId) {
            return;
        }
        if (this.isRangeDragging) {
            this.releaseRangeDragPointerCapture();
            this.commitRangeDrag(event);
            this.resetTouchGestureState();
            return;
        }
        const isTap: boolean = !this.activeTouchMoved;
        const anchor: RangeSelectionAnchor = this.getAnchorFromPointerEvent(event) ?? this.activeTouchDownAnchor;
        if (isTap) {
            if (this.touchPendingAction === "clearSelection") {
                this.clearRangeSelection(true);
            } else if (this.touchPendingAction === "setOrCommit" && anchor) {
                this.handleTouchTapRangePick(anchor);
            }
        }
        this.resetTouchGestureState();
    }

    private captureRangePointer(event: PointerEvent): void {
        const pointerCaptureElement: Element = event.currentTarget as Element;
        if (pointerCaptureElement?.setPointerCapture) {
            pointerCaptureElement.setPointerCapture(event.pointerId);
            this.rangeDragPointerCaptureElement = pointerCaptureElement;
            this.rangeDragPointerId = event.pointerId;
        }
    }

    private commitRangeDrag(event?: PointerEvent): void {
        if (!this.rangeSelection.enabled || !this.isRangeDragging || !this.dragStartAnchor) {
            return;
        }
        const anchor: RangeSelectionAnchor = (event ? this.getAnchorFromPointerEvent(event) : undefined)
            ?? this.dragCurrentAnchor
            ?? this.dragStartAnchor;
        this.isRangeDragging = false;
        this.maskDragChromePrepared = false;
        this.dragCurrentAnchor = anchor;
        const committedSelection: RangeSelectionPayload = this.createSelectionPayload("committed", this.dragStartAnchor, this.dragCurrentAnchor, false);
        // If the raw picked range contains no notes (e.g. empty-space click), keep it empty.
        // Do this before snap/padding so we don't accidentally pull in nearby notes.
        if (!this.selectionHasAnyNotes(committedSelection.normalizedStart, committedSelection.normalizedEnd)) {
            this.clearRangeSelection(true);
            return;
        }
        const paddedSelection: RangeSelectionPayload = this.applySelectionPadding(committedSelection, this.activeDragBound);
        this.dragStartAnchor = paddedSelection.normalizedStart;
        this.dragCurrentAnchor = paddedSelection.normalizedEnd;
        this.pendingTouchRangeStartAnchor = undefined;
        this.activeDragBound = "both";
        this.emitRangeHandleDragging(false);
        if (!this.selectionHasAnyNotes(this.dragStartAnchor, this.dragCurrentAnchor)) {
            this.clearRangeSelection(true);
            return;
        }
        this.renderRangeSelection();
        this.emitRangeSelection("committed", this.dragStartAnchor, this.dragCurrentAnchor, false);
    }

    private startRangeHandleDrag(
        existingSelection: RangeSelectionPayload,
        draggedBound: "start" | "end",
        anchor: RangeSelectionAnchor
    ): void {
        this.isRangeDragging = true;
        this.emitRangeHandleDragging(true);
        this.setTouchDragScrollLockEnabled(true);
        this.setTouchDragNativeScrollSuppressed(true);
        this.startTouchDragAutoScroll();
        if (draggedBound === "start") {
            // Resize start bound (keep end fixed).
            this.activeDragBound = "start";
            this.dragStartAnchor = existingSelection.normalizedEnd;
        } else {
            // Resize end bound (keep start fixed).
            this.activeDragBound = "end";
            this.dragStartAnchor = existingSelection.normalizedStart;
        }
        this.dragCurrentAnchor = anchor;
        this.renderRangeSelection();
        this.emitRangeSelection("dragging", this.dragStartAnchor, this.dragCurrentAnchor, true);
    }

    private handleTouchTapRangePick(anchor: RangeSelectionAnchor): void {
        if (!this.pendingTouchRangeStartAnchor) {
            this.pendingTouchRangeStartAnchor = anchor;
            this.hoverAnchor = anchor;
            this.renderRangeSelection();
            this.emitRangeSelection("hover", anchor, anchor, false);
            return;
        }
        this.dragStartAnchor = this.pendingTouchRangeStartAnchor;
        this.dragCurrentAnchor = anchor;
        this.pendingTouchRangeStartAnchor = undefined;
        const committedSelection: RangeSelectionPayload = this.createSelectionPayload("committed", this.dragStartAnchor, this.dragCurrentAnchor, false);
        // If the raw picked range contains no notes, don't snap-expand to nearby notes.
        if (!this.selectionHasAnyNotes(committedSelection.normalizedStart, committedSelection.normalizedEnd)) {
            this.clearRangeSelection(true);
            return;
        }
        const paddedSelection: RangeSelectionPayload = this.applySelectionPadding(committedSelection, "both");
        this.dragStartAnchor = paddedSelection.normalizedStart;
        this.dragCurrentAnchor = paddedSelection.normalizedEnd;
        if (!this.selectionHasAnyNotes(this.dragStartAnchor, this.dragCurrentAnchor)) {
            this.clearRangeSelection(true);
            return;
        }
        this.renderRangeSelection();
        this.emitRangeSelection("committed", this.dragStartAnchor, this.dragCurrentAnchor, false);
    }

    private updateTouchMoveState(event: PointerEvent): void {
        if (event.pointerId !== this.activeTouchPointerId || this.activeTouchMoved) {
            return;
        }
        const movementThresholdPx: number = 10;
        const distanceX: number = Math.abs(event.clientX - this.activeTouchStartClientX);
        const distanceY: number = Math.abs(event.clientY - this.activeTouchStartClientY);
        if (distanceX >= movementThresholdPx || distanceY >= movementThresholdPx) {
            this.activeTouchMoved = true;
            this.touchPendingAction = "none";
        }
    }

    private resetTouchGestureState(): void {
        this.stopTouchDragAutoScroll();
        this.setTouchDragScrollLockEnabled(false);
        this.setTouchDragNativeScrollSuppressed(false);
        this.activeTouchPointerId = -1;
        this.activeTouchMoved = false;
        this.activeTouchDownAnchor = undefined;
        this.activeTouchDragClientX = 0;
        this.activeTouchDragClientY = 0;
        this.touchPendingAction = "none";
    }

    private isTouchPointerEvent(event: PointerEvent): boolean {
        return event.pointerType === "touch";
    }

    private setTouchDragScrollLockEnabled(enabled: boolean): void {
        if (this.touchDragScrollLockEnabled === enabled) {
            return;
        }
        this.touchDragScrollLockEnabled = enabled;
        for (const element of this.rangeInteractionBoundElements) {
            if (enabled) {
                element.style.touchAction = "none";
            } else {
                element.style.removeProperty("touch-action");
            }
        }
    }

    private setTouchDragNativeScrollSuppressed(enabled: boolean): void {
        if (this.touchDragNativeScrollSuppressed === enabled) {
            return;
        }
        this.touchDragNativeScrollSuppressed = enabled;
        if (enabled) {
            // iOS Safari/WKWebView can keep panning unless touchmove is cancelled in capture phase.
            document.addEventListener("touchmove", this.touchMoveDuringRangeDragListener, { passive: false, capture: true });
            window.addEventListener("touchmove", this.touchMoveDuringRangeDragListener, { passive: false, capture: true });
            for (const element of this.rangeInteractionBoundElements) {
                element.addEventListener("touchmove", this.touchMoveDuringRangeDragListener, { passive: false, capture: true });
            }
        } else {
            document.removeEventListener("touchmove", this.touchMoveDuringRangeDragListener, true);
            window.removeEventListener("touchmove", this.touchMoveDuringRangeDragListener, true);
            for (const element of this.rangeInteractionBoundElements) {
                element.removeEventListener("touchmove", this.touchMoveDuringRangeDragListener, true);
            }
        }
    }

    private onTouchMoveDuringRangeDrag(event: TouchEvent): void {
        if (!this.isRangeDragging || this.activeTouchPointerId < 0) {
            return;
        }
        event.preventDefault();
        event.stopPropagation();
    }

    private emitRangeHandleDragging(isHandleDragging: boolean): void {
        if (this.isRangeHandleDragging === isHandleDragging) {
            return;
        }
        this.isRangeHandleDragging = isHandleDragging;
        if (this.rangeSelection.callbacks.onHandleDraggingChange) {
            this.rangeSelection.callbacks.onHandleDraggingChange(isHandleDragging);
        }
    }

    private updateDesktopRangeCursor(event: PointerEvent, anchor: RangeSelectionAnchor): void {
        if (this.isTouchPointerEvent(event)) {
            return;
        }
        const pointerElement: HTMLElement = event.currentTarget as HTMLElement;
        if (!pointerElement) {
            return;
        }
        const existingSelection: RangeSelectionPayload = this.getRangeSelection();
        const hoveredHandle: "start" | "end" | undefined = this.getDraggedBoundFromAnchor(anchor, existingSelection, false);
        pointerElement.style.cursor = hoveredHandle ? "pointer" : "";
    }

    private clearDesktopRangeCursor(event: PointerEvent): void {
        if (this.isTouchPointerEvent(event)) {
            return;
        }
        const pointerElement: HTMLElement = event.currentTarget as HTMLElement;
        if (!pointerElement) {
            return;
        }
        pointerElement.style.removeProperty("cursor");
    }

    private startTouchDragAutoScroll(): void {
        if (this.rangeTouchAutoScrollAnimationFrameId !== 0) {
            return;
        }
        const step: () => void = (): void => {
            this.rangeTouchAutoScrollAnimationFrameId = 0;
            if (!this.isRangeDragging || this.activeTouchPointerId < 0) {
                return;
            }
            const scrollContainer: HTMLElement | Window = this.getTouchDragScrollContainer();
            const scrollDeltaY: number = this.getTouchDragScrollDeltaY(scrollContainer, this.activeTouchDragClientY);
            if (scrollDeltaY !== 0) {
                if (this.isWindowObject(scrollContainer)) {
                    scrollContainer.scrollBy(0, scrollDeltaY);
                } else {
                    scrollContainer.scrollTop += scrollDeltaY;
                }
                const anchorFromScroll: RangeSelectionAnchor = this.getAnchorFromClientPoint(this.activeTouchDragClientX, this.activeTouchDragClientY);
                if (anchorFromScroll && this.dragStartAnchor) {
                    this.dragCurrentAnchor = anchorFromScroll;
                    this.renderRangeSelection();
                    this.emitRangeSelection("dragging", this.dragStartAnchor, this.dragCurrentAnchor, true);
                }
            }
            this.rangeTouchAutoScrollAnimationFrameId = window.requestAnimationFrame(step);
        };
        this.rangeTouchAutoScrollAnimationFrameId = window.requestAnimationFrame(step);
    }

    private stopTouchDragAutoScroll(): void {
        if (this.rangeTouchAutoScrollAnimationFrameId === 0) {
            return;
        }
        window.cancelAnimationFrame(this.rangeTouchAutoScrollAnimationFrameId);
        this.rangeTouchAutoScrollAnimationFrameId = 0;
    }

    private getTouchDragScrollContainer(): HTMLElement | Window {
        let current: HTMLElement = this.container;
        while (current && current !== document.body) {
            const styles: CSSStyleDeclaration = window.getComputedStyle(current);
            const overflowY: string = styles.overflowY;
            const isScrollable: boolean = (overflowY === "auto" || overflowY === "scroll" || overflowY === "overlay")
                && current.scrollHeight > current.clientHeight;
            if (isScrollable) {
                return current;
            }
            current = current.parentElement;
        }
        return window;
    }

    private getRangeSelectionScrollContainer(): HTMLElement | Window {
        let current: HTMLElement = this.container;
        while (current && current !== document.body) {
            const styles: CSSStyleDeclaration = window.getComputedStyle(current);
            const overflowY: string = styles.overflowY;
            const isScrollable: boolean = (overflowY === "auto" || overflowY === "scroll" || overflowY === "overlay")
                && current.scrollHeight > current.clientHeight;
            if (isScrollable) {
                return current;
            }
            current = current.parentElement;
        }
        return window;
    }

    private getTouchDragScrollDeltaY(scrollContainer: HTMLElement | Window, clientY: number): number {
        const edgeThresholdPx: number = this.getRangeSelectionSnapPaddingPx();
        const minSpeedPxPerFrame: number = 2;
        const maxSpeedPxPerFrame: number = 10;
        let top: number;
        let bottom: number;
        if (this.isWindowObject(scrollContainer)) {
            top = 0;
            bottom = window.innerHeight;
        } else {
            const rect: DOMRect = scrollContainer.getBoundingClientRect();
            top = rect.top;
            bottom = rect.bottom;
        }
        if (clientY < top + edgeThresholdPx) {
            const ratio: number = Math.max(0, (top + edgeThresholdPx - clientY) / edgeThresholdPx);
            const speed: number = minSpeedPxPerFrame + (maxSpeedPxPerFrame - minSpeedPxPerFrame) * ratio;
            return -Math.round(speed);
        }
        if (clientY > bottom - edgeThresholdPx) {
            const ratio: number = Math.max(0, (clientY - (bottom - edgeThresholdPx)) / edgeThresholdPx);
            const speed: number = minSpeedPxPerFrame + (maxSpeedPxPerFrame - minSpeedPxPerFrame) * ratio;
            return Math.round(speed);
        }
        return 0;
    }

    private isWindowObject(target: HTMLElement | Window): target is Window {
        return target === window;
    }

    private getAnchorFromPointerEvent(event: PointerEvent): RangeSelectionAnchor {
        return this.getAnchorFromClientPoint(event.clientX, event.clientY);
    }

    private getAnchorFromClientPoint(clientX: number, clientY: number): RangeSelectionAnchor {
        if (!this.graphic) {
            return undefined;
        }
        const domPoint: PointF2D = new PointF2D(clientX, clientY);
        const svgPoint: PointF2D = this.graphic.domToSvg(domPoint);
        if (!svgPoint) {
            return undefined;
        }
        const osmdPoint: PointF2D = this.graphic.svgToOsmd(svgPoint);
        if (!osmdPoint) {
            return undefined;
        }
        const system: MusicSystem = this.findSystemAtPosition(osmdPoint);
        if (!system) {
            return undefined;
        }
        const leftBoundaryX: number = system.GetLeftBorderAbsoluteXPosition();
        const rightBoundaryX: number = system.GetRightBorderAbsoluteXPosition();
        const zoomScale: number = Math.max(0.0001, this.zoom * 10.0);
        const preStartPaddingPx: number = this.getRangeSelectionPreStartPaddingPx();
        const preStartPaddingOsmd: number = preStartPaddingPx / zoomScale;
        const x: number = Math.min(
            rightBoundaryX,
            Math.max(leftBoundaryX - preStartPaddingOsmd, osmdPoint.x)
        );
        const xPx: number = x * zoomScale;
        const startTime: Fraction = system.GetSystemsFirstTimeStamp();
        const endTime: Fraction = system.GetSystemsLastTimeStamp();
        const totalWidthOsmd: number = Math.max(0.0001, rightBoundaryX - leftBoundaryX);
        const ratio: number = (x - leftBoundaryX) / totalWidthOsmd;
        const clampedRatio: number = Math.max(-(preStartPaddingOsmd / totalWidthOsmd), Math.min(1, ratio));
        const timestampReal: number = startTime.RealValue + (endTime.RealValue - startTime.RealValue) * clampedRatio;
        const timestamp: Fraction = new Fraction(timestampReal, 1);
        if (!timestamp) {
            return undefined;
        }
        const staffEntry: GraphicalStaffEntry = this.graphic.GetNearestStaffEntry(new PointF2D(x, osmdPoint.y));
        const lineBounds: { yPx: number, heightPx: number } = this.getSystemVerticalBoundsInPixels(system);
        return {
            timestamp,
            timestampReal,
            measureIndex: staffEntry?.parentMeasure?.parentSourceMeasure?.measureListIndex
                ?? system.GraphicalMeasures[0]?.[0]?.parentSourceMeasure?.measureListIndex
                ?? 0,
            systemIndex: this.getSystemIndex(system),
            pageNumber: system.Parent?.PageNumber ?? 1,
            staffIndex: staffEntry?.sourceStaffEntry?.ParentStaff?.idInMusicSheet ?? system.StaffLines[0]?.ParentStaff?.idInMusicSheet ?? 0,
            x,
            xPx,
            yPx: lineBounds.yPx,
            heightPx: lineBounds.heightPx
        };
    }

    private createAnchorFromTimestamp(timestamp: Fraction): RangeSelectionAnchor {
        if (!this.graphic) {
            return undefined;
        }
        const result: [number, MusicSystem] = this.graphic.calculateXPositionFromTimestamp(timestamp);
        if (!result || !result[1]) {
            return undefined;
        }
        const x: number = result[0];
        const system: MusicSystem = result[1];
        const systemMeasure: GraphicalMeasure = system.GraphicalMeasures[0]?.[0];
        const lineBounds: { yPx: number, heightPx: number } = this.getSystemVerticalBoundsInPixels(system);
        return {
            timestamp,
            timestampReal: timestamp.RealValue,
            measureIndex: systemMeasure?.parentSourceMeasure?.measureListIndex ?? 0,
            systemIndex: this.getSystemIndex(system),
            pageNumber: system.Parent?.PageNumber ?? 1,
            staffIndex: system.StaffLines[0]?.ParentStaff?.idInMusicSheet ?? 0,
            x,
            xPx: x * this.zoom * 10.0,
            yPx: lineBounds.yPx,
            heightPx: lineBounds.heightPx
        };
    }

    private getSystemVerticalBoundsInPixels(system: MusicSystem): { yPx: number, heightPx: number } {
        const firstStaffLine: MusicSystem["StaffLines"][number] = system.StaffLines[0];
        const lastStaffLine: MusicSystem["StaffLines"][number] = system.StaffLines[system.StaffLines.length - 1];
        const y: number = system.PositionAndShape.AbsolutePosition.y + firstStaffLine.PositionAndShape.RelativePosition.y;
        const bottomY: number = system.PositionAndShape.AbsolutePosition.y + lastStaffLine.PositionAndShape.RelativePosition.y + lastStaffLine.StaffHeight;
        const scale: number = this.zoom * 10.0;
        return {
            yPx: y * scale,
            heightPx: (bottomY - y) * scale
        };
    }

    private getSystemHorizontalBoundsInPixels(system: MusicSystem): { leftPx: number, rightPx: number } {
        const scale: number = this.zoom * 10.0;
        return {
            leftPx: system.GetLeftBorderAbsoluteXPosition() * scale,
            rightPx: system.GetRightBorderAbsoluteXPosition() * scale
        };
    }

    /**
     * Full visual width for outside-range masks. Unlike {@link getSystemHorizontalBoundsInPixels},
     * the left edge is the staff-line origin (before clefs / key / time signatures) so gray-out
     * covers the whole system, not just the area to the right of begin instructions.
     */
    private getSystemMaskHorizontalBoundsInPixels(system: MusicSystem): { leftPx: number, rightPx: number } {
        const scale: number = this.zoom * 10.0;
        let leftBorderX: number = Number.POSITIVE_INFINITY;
        for (const staffLine of system.StaffLines ?? []) {
            const staffStartX: number = staffLine?.PositionAndShape?.AbsolutePosition?.x;
            if (Number.isFinite(staffStartX)) {
                leftBorderX = Math.min(leftBorderX, staffStartX);
            }
        }
        if (!Number.isFinite(leftBorderX)) {
            leftBorderX = system.GetLeftBorderAbsoluteXPosition();
        }
        return {
            leftPx: leftBorderX * scale,
            rightPx: system.GetRightBorderAbsoluteXPosition() * scale
        };
    }

    private getSystemIndex(targetSystem: MusicSystem): number {
        let index: number = 0;
        for (const page of this.graphic.MusicPages) {
            for (const system of page.MusicSystems) {
                if (system === targetSystem) {
                    return index;
                }
                index++;
            }
        }
        return -1;
    }

    private findSystemAtPosition(position: PointF2D): MusicSystem {
        if (!this.graphic) {
            return undefined;
        }
        let closestSystem: MusicSystem = undefined;
        let closestDistance: number = Number.MAX_VALUE;
        for (const page of this.graphic.MusicPages) {
            for (const system of page.MusicSystems) {
                const left: number = system.GetLeftBorderAbsoluteXPosition();
                const right: number = system.GetRightBorderAbsoluteXPosition();
                const firstStaffLine: MusicSystem["StaffLines"][number] = system.StaffLines[0];
                const lastStaffLine: MusicSystem["StaffLines"][number] = system.StaffLines[system.StaffLines.length - 1];
                if (!firstStaffLine || !lastStaffLine) {
                    continue;
                }
                const top: number = system.PositionAndShape.AbsolutePosition.y + firstStaffLine.PositionAndShape.RelativePosition.y;
                const bottom: number = system.PositionAndShape.AbsolutePosition.y
                    + lastStaffLine.PositionAndShape.RelativePosition.y + lastStaffLine.StaffHeight;
                if (position.x >= left && position.x <= right && position.y >= top && position.y <= bottom) {
                    return system;
                }
                const dx: number = position.x < left ? left - position.x : (position.x > right ? position.x - right : 0);
                const dy: number = position.y < top ? top - position.y : (position.y > bottom ? position.y - bottom : 0);
                const distance: number = dx + dy;
                if (distance < closestDistance) {
                    closestDistance = distance;
                    closestSystem = system;
                }
            }
        }
        return closestSystem;
    }

    private isAnchorInsideSelection(anchor: RangeSelectionAnchor, selection: RangeSelectionPayload): boolean {
        if (!anchor || !selection) {
            return false;
        }
        return anchor.timestampReal >= selection.normalizedStart.timestampReal
            && anchor.timestampReal <= selection.normalizedEnd.timestampReal;
    }

    private getDraggedBoundFromAnchor(
        anchor: RangeSelectionAnchor,
        selection: RangeSelectionPayload,
        isTouchInteraction: boolean = false
    ): "start" | "end" | undefined {
        if (!anchor || !selection) {
            return undefined;
        }
        const zoomValue: number = Math.max(0.25, Number.isFinite(this.zoom) ? this.zoom : 1);
        const configuredLineWidthPx: number = this.rangeSelection.options.lineWidthPx ?? 12;
        // Desktop precision should improve as zoom increases (smaller hit target),
        // while touch remains forgiving regardless of zoom.
        const lineHitTolerancePx: number = isTouchInteraction
            ? Math.max(24, Math.min(38, configuredLineWidthPx * 2.2))
            : Math.max(4, Math.min(10, configuredLineWidthPx / Math.sqrt(zoomValue)));
        const matchesStartSystem: boolean = anchor.systemIndex === selection.normalizedStart.systemIndex;
        const matchesEndSystem: boolean = anchor.systemIndex === selection.normalizedEnd.systemIndex;
        const startDistancePx: number = Math.abs(anchor.xPx - selection.normalizedStart.xPx);
        const endDistancePx: number = Math.abs(anchor.xPx - selection.normalizedEnd.xPx);
        const nearStartLine: boolean = matchesStartSystem && startDistancePx <= lineHitTolerancePx;
        const nearEndLine: boolean = matchesEndSystem && endDistancePx <= lineHitTolerancePx;
        if (!nearStartLine && !nearEndLine) {
            return undefined;
        }
        if (nearStartLine && nearEndLine) {
            return startDistancePx <= endDistancePx ? "start" : "end";
        }
        return nearStartLine ? "start" : "end";
    }

    private applySelectionPadding(selection: RangeSelectionPayload, movedBound: "start" | "end" | "both" = "both"): RangeSelectionPayload {
        if (!selection) {
            return selection;
        }
        // Snap-to-notes: align selection boundaries to nearest note positions
        if (this.rangeSelection.options.snapToNotes && this.graphic) {
            const snappedStart: RangeSelectionAnchor = movedBound === "end"
                ? selection.normalizedStart
                : this.snapAnchorToNearestNote(selection.normalizedStart, "start");
            const snappedEnd: RangeSelectionAnchor = movedBound === "start"
                ? selection.normalizedEnd
                : this.snapAnchorToNearestNote(selection.normalizedEnd, "end");
            return this.createSelectionPayload(selection.phase, snappedStart, snappedEnd, selection.isDragging);
        }
        const paddingPx: number = this.rangeSelection.options.applyPaddingPx ?? 0;
        if (!Number.isFinite(paddingPx) || paddingPx <= 0) {
            return selection;
        }
        const paddedStart: RangeSelectionAnchor = movedBound === "end"
            ? selection.normalizedStart
            : this.shiftAnchorX(selection.normalizedStart, -paddingPx);
        const paddedEnd: RangeSelectionAnchor = movedBound === "start"
            ? selection.normalizedEnd
            : this.shiftAnchorX(selection.normalizedEnd, paddingPx);
        return this.createSelectionPayload(selection.phase, paddedStart, paddedEnd, selection.isDragging);
    }

    /**
     * Snap an anchor to the nearest GraphicalStaffEntry note position.
     * For "start" bound: snap to the first note at or after the anchor's timestamp.
     * For "end" bound: snap to the last note at or before the anchor's timestamp.
     */
    private snapAnchorToNearestNote(anchor: RangeSelectionAnchor, bound: "start" | "end"): RangeSelectionAnchor {
        if (!anchor || !this.graphic) {
            return anchor;
        }
        const scale: number = this.zoom * 10.0;
        let bestEntry: GraphicalStaffEntry = undefined;
        let bestSystemIndex: number = -1;
        let bestXPx: number = bound === "start" ? Number.MAX_VALUE : Number.MIN_VALUE;
        let systemIndex: number = 0;
        for (const page of this.graphic.MusicPages) {
            for (const musicSystem of page.MusicSystems) {
                for (const staffLine of musicSystem.StaffLines) {
                    for (const measure of staffLine.Measures) {
                        for (const entry of measure.staffEntries) {
                            if (!this.entryHasPlayableNotes(entry) || !this.isStaffEntryVisibleForRangeSnap(entry)) {
                                continue;
                            }
                            const candidateX: number = entry.PositionAndShape?.AbsolutePosition?.x;
                            if (typeof candidateX !== "number" || !Number.isFinite(candidateX)) {
                                continue;
                            }
                            const candidateXPx: number = candidateX * scale;
                            // Use xPx + systemIndex matching (same criteria as gray-out logic)
                            if (bound === "start") {
                                // Find first note at or after anchor xPx (in same system) or in later system
                                const isAfterAnchor: boolean = systemIndex > anchor.systemIndex
                                    || (systemIndex === anchor.systemIndex && candidateXPx >= anchor.xPx - 1);
                                const isBetterThanBest: boolean = bestEntry === undefined
                                    || systemIndex < bestSystemIndex
                                    || (systemIndex === bestSystemIndex && candidateXPx < bestXPx);
                                if (isAfterAnchor && isBetterThanBest) {
                                    bestEntry = entry;
                                    bestSystemIndex = systemIndex;
                                    bestXPx = candidateXPx;
                                }
                            } else {
                                // Find last note at or before anchor xPx (in same system) or in earlier system
                                const isBeforeAnchor: boolean = systemIndex < anchor.systemIndex
                                    || (systemIndex === anchor.systemIndex && candidateXPx <= anchor.xPx + 1);
                                const isBetterThanBest: boolean = bestEntry === undefined
                                    || systemIndex > bestSystemIndex
                                    || (systemIndex === bestSystemIndex && candidateXPx > bestXPx);
                                if (isBeforeAnchor && isBetterThanBest) {
                                    bestEntry = entry;
                                    bestSystemIndex = systemIndex;
                                    bestXPx = candidateXPx;
                                }
                            }
                        }
                    }
                }
                systemIndex++;
            }
        }
        if (!bestEntry) {
            return anchor;
        }
        const entryX: number = bestEntry.PositionAndShape?.AbsolutePosition?.x ?? anchor.x;
        const entryXPx: number = entryX * scale;
        const widthCompensationPx: number = this.getSnapWidthCompensationPxForEntry(bestEntry, bound, scale);
        const snapPaddingPx: number = this.getRangeSelectionSnapPaddingPx();
        const rawSnappedXPx: number = bound === "start"
            ? entryXPx - snapPaddingPx - widthCompensationPx
            : entryXPx + snapPaddingPx + widthCompensationPx;
        const snappedXPx: number = this.clampSnapBoundaryToAdjacentEntryMidpoint(bestEntry, bound, rawSnappedXPx, scale);
        const snappedX: number = snappedXPx / scale;
        const entryTimestamp: number = bestEntry.getAbsoluteTimestamp()?.RealValue ?? anchor.timestampReal;
        const system: MusicSystem = bestEntry.parentMeasure?.ParentMusicSystem
            ?? this.findSystemByIndex(anchor.systemIndex);
        const lineBounds: { yPx: number, heightPx: number } = system
            ? this.getSystemVerticalBoundsInPixels(system)
            : { yPx: anchor.yPx, heightPx: anchor.heightPx };
        return {
            timestamp: new Fraction(entryTimestamp, 1),
            timestampReal: entryTimestamp,
            measureIndex: bestEntry.parentMeasure?.parentSourceMeasure?.measureListIndex ?? anchor.measureIndex,
            systemIndex: system ? this.getSystemIndex(system) : anchor.systemIndex,
            pageNumber: system?.Parent?.PageNumber ?? anchor.pageNumber,
            staffIndex: bestEntry.sourceStaffEntry?.ParentStaff?.idInMusicSheet ?? anchor.staffIndex,
            x: snappedX,
            xPx: snappedXPx,
            selectionX: entryX,
            selectionXPx: entryXPx,
            yPx: lineBounds.yPx,
            heightPx: lineBounds.heightPx,
        };
    }

    private getRangeSelectionSnapPaddingPx(): number {
        // Keep this aligned with touch auto-scroll edge threshold and
        // scale with zoom so snap spacing remains musically consistent.
        const zoomValue: number = Math.max(0.25, Number.isFinite(this.zoom) ? this.zoom : 1);
        const zoomScaledPaddingPx: number = 18 * Math.sqrt(zoomValue);
        return Math.max(18, Math.min(28, zoomScaledPaddingPx));
    }

    private entryHasPlayableNotes(entry: GraphicalStaffEntry): boolean {
        return entry?.sourceStaffEntry?.VoiceEntries?.some(
            (voiceEntry: any) => voiceEntry.Notes?.some((note: any) => !note.isRest())
        ) ?? false;
    }

    private clampSnapBoundaryToAdjacentEntryMidpoint(
        entry: GraphicalStaffEntry,
        bound: "start" | "end",
        snappedXPx: number,
        scale: number
    ): number {
        if (!entry || !Number.isFinite(snappedXPx) || !Number.isFinite(scale)) {
            return snappedXPx;
        }
        const neighborBoundsXPx: { leftPx: number, rightPx: number } = this.getAdjacentPlayableEntryBoundsXPx(
            entry,
            bound === "start" ? "previous" : "next",
            scale
        );
        const allowMeasureBoundarySnap: boolean = this.isMeasureBoundarySnapAllowed(entry, bound);
        const epsilonPx: number = 0.5;
        const measureBounds: { leftPx: number, rightPx: number } = this.getMeasureHorizontalBoundsInPixels(
            entry?.parentMeasure,
            entry?.parentMeasure?.ParentMusicSystem
        );

        if (bound === "start" && this.isFirstVisiblePlayableEntryInSystem(entry, scale)) {
            const system: MusicSystem = entry?.parentMeasure?.ParentMusicSystem;
            if (system) {
                // For the first visible playable note in a system, allow snapping to the
                // pre-start area left of the opening clef/key/time block so selection can
                // include beat-0 notes and feel anchored to the visual system start.
                const systemBounds: { leftPx: number, rightPx: number } = this.getSystemHorizontalBoundsInPixels(system);
                const leftVisualStartPx: number = Number.isFinite(measureBounds.leftPx)
                    ? Math.min(systemBounds.leftPx, measureBounds.leftPx)
                    : systemBounds.leftPx;
                return leftVisualStartPx - this.getRangeSelectionPreStartPaddingPx();
            }
        }

        // Rule: if this is the first/last visible playable timestamp in the measure,
        // allow measure boundary snapping directly.
        if (allowMeasureBoundarySnap) {
            if (bound === "start") {
                return Number.isFinite(measureBounds.leftPx)
                    ? measureBounds.leftPx
                    : snappedXPx;
            }
            return Number.isFinite(measureBounds.rightPx)
                ? measureBounds.rightPx
                : snappedXPx;
        }

        if (bound === "start") {
            if (neighborBoundsXPx && Number.isFinite(neighborBoundsXPx.rightPx)) {
                // Prefer snapping almost to the previous visible note edge.
                // This keeps the range tight without including that note.
                const entryBoundsXPx: { leftPx: number, rightPx: number } = this.getEntryPlayableBoundsXPx(entry, scale);
                const maxPreviousSnapDistancePx: number = 32;
                const previousNoteSnapXPx: number = neighborBoundsXPx.rightPx + epsilonPx;
                if (entryBoundsXPx && Number.isFinite(entryBoundsXPx.leftPx)) {
                    // Cap how far back start-handle can jump when snapping to previous note.
                    const furthestBackAllowedXPx: number = entryBoundsXPx.leftPx - maxPreviousSnapDistancePx;
                    return Math.max(furthestBackAllowedXPx, previousNoteSnapXPx);
                }
                return previousNoteSnapXPx;
            }
            // No adjacent visible note on this side: fall back to snap padding.
            return snappedXPx;
        }

        if (neighborBoundsXPx && Number.isFinite(neighborBoundsXPx.leftPx)) {
            // Prefer snapping almost to the next visible note edge.
            // This keeps the range tight without including that note.
            return neighborBoundsXPx.leftPx - epsilonPx;
        }
        // No adjacent visible note on this side: fall back to snap padding.
        return snappedXPx;
    }

    private isMeasureBoundarySnapAllowed(
        entry: GraphicalStaffEntry,
        bound: "start" | "end"
    ): boolean {
        const measure: GraphicalMeasure = entry?.parentMeasure;
        if (!measure) {
            return false;
        }
        const entryTimestamp: number = entry?.getAbsoluteTimestamp()?.RealValue;
        if (!Number.isFinite(entryTimestamp)) {
            return false;
        }
        const targetSystem: MusicSystem = measure?.ParentMusicSystem;
        const targetMeasureIndex: number = measure?.parentSourceMeasure?.measureListIndex;
        if (!targetSystem || !Number.isFinite(targetMeasureIndex)) {
            return false;
        }
        const scale: number = this.zoom * 10.0;
        if (!Number.isFinite(scale)) {
            return false;
        }
        type EntryBoundarySnapshot = { timestamp: number, leftPx: number, rightPx: number };
        const boundarySnapshots: EntryBoundarySnapshot[] = [];
        let minTimestamp: number = Number.POSITIVE_INFINITY;
        let maxTimestamp: number = Number.NEGATIVE_INFINITY;
        for (const staffLine of targetSystem.StaffLines ?? []) {
            for (const siblingMeasure of staffLine?.Measures ?? []) {
                if (siblingMeasure?.parentSourceMeasure?.measureListIndex !== targetMeasureIndex) {
                    continue;
                }
                for (const candidateEntry of siblingMeasure?.staffEntries ?? []) {
                    if (!candidateEntry
                        || !this.entryHasPlayableNotes(candidateEntry)
                        || !this.isStaffEntryVisibleForRangeSnap(candidateEntry)) {
                        continue;
                    }
                    const candidateTimestamp: number = candidateEntry.getAbsoluteTimestamp()?.RealValue;
                    if (!Number.isFinite(candidateTimestamp)) {
                        continue;
                    }
                    const candidateBoundsXPx: { leftPx: number, rightPx: number } = this.getEntryPlayableBoundsXPx(candidateEntry, scale);
                    if (!candidateBoundsXPx
                        || !Number.isFinite(candidateBoundsXPx.leftPx)
                        || !Number.isFinite(candidateBoundsXPx.rightPx)) {
                        continue;
                    }
                    minTimestamp = Math.min(minTimestamp, candidateTimestamp);
                    maxTimestamp = Math.max(maxTimestamp, candidateTimestamp);
                    boundarySnapshots.push({
                        timestamp: candidateTimestamp,
                        leftPx: candidateBoundsXPx.leftPx,
                        rightPx: candidateBoundsXPx.rightPx
                    });
                }
            }
        }
        if (!Number.isFinite(minTimestamp) || !Number.isFinite(maxTimestamp) || boundarySnapshots.length === 0) {
            return false;
        }
        const entryBoundsXPx: { leftPx: number, rightPx: number } = this.getEntryPlayableBoundsXPx(entry, scale);
        if (!entryBoundsXPx || !Number.isFinite(entryBoundsXPx.leftPx) || !Number.isFinite(entryBoundsXPx.rightPx)) {
            return false;
        }
        const epsilon: number = Fraction.FloatInaccuracyTolerance;
        const xEpsilonPx: number = 0.5;
        if (bound === "start") {
            if (entryTimestamp > minTimestamp + epsilon) {
                return false;
            }
            let minLeftAtMinTimestamp: number = Number.POSITIVE_INFINITY;
            for (const snapshot of boundarySnapshots) {
                if (snapshot.timestamp <= minTimestamp + epsilon) {
                    minLeftAtMinTimestamp = Math.min(minLeftAtMinTimestamp, snapshot.leftPx);
                }
            }
            return Number.isFinite(minLeftAtMinTimestamp)
                && entryBoundsXPx.leftPx <= minLeftAtMinTimestamp + xEpsilonPx;
        }
        if (entryTimestamp < maxTimestamp - epsilon) {
            return false;
        }
        let maxRightAtMaxTimestamp: number = Number.NEGATIVE_INFINITY;
        for (const snapshot of boundarySnapshots) {
            if (snapshot.timestamp >= maxTimestamp - epsilon) {
                maxRightAtMaxTimestamp = Math.max(maxRightAtMaxTimestamp, snapshot.rightPx);
            }
        }
        return Number.isFinite(maxRightAtMaxTimestamp)
            && entryBoundsXPx.rightPx >= maxRightAtMaxTimestamp - xEpsilonPx;
    }

    private isFirstVisiblePlayableEntryInSystem(entry: GraphicalStaffEntry, scale: number): boolean {
        const system: MusicSystem = entry?.parentMeasure?.ParentMusicSystem;
        if (!system || !Number.isFinite(scale)) {
            return false;
        }
        const entryTimestamp: number = entry?.getAbsoluteTimestamp()?.RealValue;
        const entryBoundsXPx: { leftPx: number, rightPx: number } = this.getEntryPlayableBoundsXPx(entry, scale);
        if (!Number.isFinite(entryTimestamp) || !entryBoundsXPx || !Number.isFinite(entryBoundsXPx.leftPx)) {
            return false;
        }
        let minTimestamp: number = Number.POSITIVE_INFINITY;
        let minLeftAtMinTimestamp: number = Number.POSITIVE_INFINITY;
        for (const staffLine of system.StaffLines ?? []) {
            for (const measure of staffLine?.Measures ?? []) {
                for (const candidateEntry of measure?.staffEntries ?? []) {
                    if (!candidateEntry
                        || !this.entryHasPlayableNotes(candidateEntry)
                        || !this.isStaffEntryVisibleForRangeSnap(candidateEntry)) {
                        continue;
                    }
                    const candidateTimestamp: number = candidateEntry.getAbsoluteTimestamp()?.RealValue;
                    const candidateBoundsXPx: { leftPx: number, rightPx: number } = this.getEntryPlayableBoundsXPx(candidateEntry, scale);
                    if (!Number.isFinite(candidateTimestamp) || !candidateBoundsXPx || !Number.isFinite(candidateBoundsXPx.leftPx)) {
                        continue;
                    }
                    if (candidateTimestamp < minTimestamp - Fraction.FloatInaccuracyTolerance) {
                        minTimestamp = candidateTimestamp;
                        minLeftAtMinTimestamp = candidateBoundsXPx.leftPx;
                    } else if (Math.abs(candidateTimestamp - minTimestamp) <= Fraction.FloatInaccuracyTolerance) {
                        minLeftAtMinTimestamp = Math.min(minLeftAtMinTimestamp, candidateBoundsXPx.leftPx);
                    }
                }
            }
        }
        if (!Number.isFinite(minTimestamp) || !Number.isFinite(minLeftAtMinTimestamp)) {
            return false;
        }
        const xEpsilonPx: number = 0.5;
        return entryTimestamp <= minTimestamp + Fraction.FloatInaccuracyTolerance
            && entryBoundsXPx.leftPx <= minLeftAtMinTimestamp + xEpsilonPx;
    }

    private isStaffEntryVisibleForRangeSnap(entry: GraphicalStaffEntry): boolean {
        if (!entry) {
            return false;
        }
        const parentStaff: any = entry?.sourceStaffEntry?.ParentStaff;
        const isStaffVisible: boolean = typeof parentStaff?.isVisible === "function"
            ? parentStaff.isVisible()
            : parentStaff?.Visible !== false;
        if (!isStaffVisible) {
            return false;
        }
        const parentMeasure: any = entry?.parentMeasure;
        return typeof parentMeasure?.isVisible === "function"
            ? parentMeasure.isVisible()
            : true;
    }

    private getAdjacentPlayableEntryBoundsXPx(
        entry: GraphicalStaffEntry,
        direction: "previous" | "next",
        scale: number
    ): { leftPx: number, rightPx: number } {
        const measure: GraphicalMeasure = entry?.parentMeasure;
        if (!measure || !Number.isFinite(scale)) {
            return undefined;
        }
        const entryCenterXPx: number = (entry.PositionAndShape?.AbsolutePosition?.x ?? NaN) * scale;
        if (!Number.isFinite(entryCenterXPx)) {
            return undefined;
        }
        const targetSystem: MusicSystem = measure?.ParentMusicSystem;
        const targetMeasureIndex: number = measure?.parentSourceMeasure?.measureListIndex;
        if (!targetSystem || !Number.isFinite(targetMeasureIndex)) {
            return undefined;
        }
        let bestCandidateEntry: GraphicalStaffEntry = undefined;
        let candidateCenterXPx: number = direction === "previous" ? Number.NEGATIVE_INFINITY : Number.POSITIVE_INFINITY;
        for (const staffLine of targetSystem.StaffLines ?? []) {
            for (const siblingMeasure of staffLine?.Measures ?? []) {
                if (siblingMeasure?.parentSourceMeasure?.measureListIndex !== targetMeasureIndex) {
                    continue;
                }
                for (const candidateEntry of siblingMeasure?.staffEntries ?? []) {
                    if (!candidateEntry
                        || candidateEntry === entry
                        || !this.entryHasPlayableNotes(candidateEntry)
                        || !this.isStaffEntryVisibleForRangeSnap(candidateEntry)) {
                        continue;
                    }
                    // Adjacent-note clamping should consider all visible entries in this
                    // source measure across all currently shown staves.
                    const centerXPx: number = (candidateEntry.PositionAndShape?.AbsolutePosition?.x ?? NaN) * scale;
                    if (!Number.isFinite(centerXPx)) {
                        continue;
                    }
                    if (direction === "previous") {
                        if (centerXPx < entryCenterXPx && centerXPx > candidateCenterXPx) {
                            candidateCenterXPx = centerXPx;
                            bestCandidateEntry = candidateEntry;
                        }
                    } else if (centerXPx > entryCenterXPx && centerXPx < candidateCenterXPx) {
                        candidateCenterXPx = centerXPx;
                        bestCandidateEntry = candidateEntry;
                    }
                }
            }
        }
        if (direction === "previous" && !Number.isFinite(candidateCenterXPx)) {
            return undefined;
        }
        if (direction === "next" && !Number.isFinite(candidateCenterXPx)) {
            return undefined;
        }
        return this.getEntryPlayableBoundsXPx(bestCandidateEntry, scale);
    }

    private getEntryPlayableBoundsXPx(
        entry: GraphicalStaffEntry,
        scale: number
    ): { leftPx: number, rightPx: number } {
        if (!entry || !Number.isFinite(scale)) {
            return undefined;
        }
        const entryXPx: number = (entry.PositionAndShape?.AbsolutePosition?.x ?? NaN) * scale;
        if (!Number.isFinite(entryXPx)) {
            return undefined;
        }
        let minLeftXPx: number = Number.POSITIVE_INFINITY;
        let maxRightXPx: number = Number.NEGATIVE_INFINITY;
        for (const graphicalVoiceEntry of entry.graphicalVoiceEntries ?? []) {
            for (const graphicalNote of graphicalVoiceEntry?.notes ?? []) {
                const sourceNote: any = graphicalNote?.sourceNote;
                if (!sourceNote || sourceNote.isRest()) {
                    continue;
                }
                const noteX: number = graphicalNote.PositionAndShape?.AbsolutePosition?.x;
                if (!Number.isFinite(noteX)) {
                    continue;
                }
                const noteXPx: number = noteX * scale;
                const noteLeftXPx: number = noteXPx + ((graphicalNote.PositionAndShape?.BorderLeft ?? 0) * scale);
                const noteRightXPx: number = noteXPx + ((graphicalNote.PositionAndShape?.BorderRight ?? 0) * scale);
                minLeftXPx = Math.min(minLeftXPx, noteLeftXPx);
                maxRightXPx = Math.max(maxRightXPx, noteRightXPx);
            }
        }
        if (!Number.isFinite(minLeftXPx) || !Number.isFinite(maxRightXPx)) {
            return { leftPx: entryXPx, rightPx: entryXPx };
        }
        return { leftPx: minLeftXPx, rightPx: maxRightXPx };
    }

    private getSnapWidthCompensationPxForEntry(
        entry: GraphicalStaffEntry,
        bound: "start" | "end",
        scale: number
    ): number {
        const entryBoundsXPx: { leftPx: number, rightPx: number } = this.getEntryPlayableBoundsXPx(entry, scale);
        if (!entryBoundsXPx) {
            return 0;
        }
        const entryXPx: number = (entry.PositionAndShape?.AbsolutePosition?.x ?? NaN) * scale;
        if (!Number.isFinite(entryXPx)) {
            return 0;
        }
        return bound === "start"
            ? Math.max(0, entryXPx - entryBoundsXPx.leftPx)
            : Math.max(0, entryBoundsXPx.rightPx - entryXPx);
    }

    private getMeasureHorizontalBoundsInPixels(
        measure: GraphicalMeasure,
        fallbackSystem?: MusicSystem
    ): { leftPx: number, rightPx: number } {
        const scale: number = this.zoom * 10.0;
        const fallbackBounds: { leftPx: number, rightPx: number } = fallbackSystem
            ? this.getSystemHorizontalBoundsInPixels(fallbackSystem)
            : { leftPx: Number.NEGATIVE_INFINITY, rightPx: Number.POSITIVE_INFINITY };
        if (!measure || !Number.isFinite(scale)) {
            return fallbackBounds;
        }

        const absoluteX: number = measure.PositionAndShape?.AbsolutePosition?.x;
        const borderLeft: number = measure.PositionAndShape?.BorderLeft ?? 0;
        const borderRight: number = measure.PositionAndShape?.BorderRight ?? 0;
        const leftPx: number = (absoluteX + borderLeft) * scale;
        const rightPx: number = (absoluteX + borderRight) * scale;
        if (!Number.isFinite(leftPx) || !Number.isFinite(rightPx)) {
            return fallbackBounds;
        }
        const normalizedLeftPx: number = Math.min(leftPx, rightPx);
        const normalizedRightPx: number = Math.max(leftPx, rightPx);
        return {
            leftPx: Number.isFinite(fallbackBounds.leftPx) ? Math.max(normalizedLeftPx, fallbackBounds.leftPx) : normalizedLeftPx,
            rightPx: Number.isFinite(fallbackBounds.rightPx) ? Math.min(normalizedRightPx, fallbackBounds.rightPx) : normalizedRightPx
        };
    }

    private shiftAnchorX(anchor: RangeSelectionAnchor, deltaPx: number): RangeSelectionAnchor {
        if (!anchor || !this.graphic) {
            return anchor;
        }
        const system: MusicSystem = this.findSystemByIndex(anchor.systemIndex);
        if (!system) {
            return anchor;
        }
        const horizontalBounds: { leftPx: number, rightPx: number } = this.getSystemHorizontalBoundsInPixels(system);
        const preStartPaddingPx: number = this.getRangeSelectionPreStartPaddingPx();
        const shiftedXPx: number = Math.min(
            horizontalBounds.rightPx,
            Math.max(horizontalBounds.leftPx - preStartPaddingPx, anchor.xPx + deltaPx)
        );
        const startTime: Fraction = system.GetSystemsFirstTimeStamp();
        const endTime: Fraction = system.GetSystemsLastTimeStamp();
        const widthPx: number = Math.max(1, horizontalBounds.rightPx - horizontalBounds.leftPx);
        const ratio: number = (shiftedXPx - horizontalBounds.leftPx) / widthPx;
        const timestampReal: number = startTime.RealValue + (endTime.RealValue - startTime.RealValue) * ratio;
        return {
            ...anchor,
            timestamp: new Fraction(timestampReal, 1),
            timestampReal,
            xPx: shiftedXPx,
            x: shiftedXPx / (this.zoom * 10.0)
        };
    }

    private findSystemByIndex(targetIndex: number): MusicSystem {
        if (!this.graphic || targetIndex < 0) {
            return undefined;
        }
        let index: number = 0;
        for (const page of this.graphic.MusicPages) {
            for (const system of page.MusicSystems) {
                if (index === targetIndex) {
                    return system;
                }
                index++;
            }
        }
        return undefined;
    }

    private createSelectionPayload(
        phase: "hover" | "dragging" | "committed" | "cleared",
        start: RangeSelectionAnchor,
        end: RangeSelectionAnchor,
        isDragging: boolean
    ): RangeSelectionPayload {
        const normalizedStart: RangeSelectionAnchor = start.timestampReal <= end.timestampReal ? start : end;
        const normalizedEnd: RangeSelectionAnchor = normalizedStart === start ? end : start;
        const direction: RangeSelectionDirection = start.timestampReal <= end.timestampReal ? "forward" : "backward";
        return {
            phase,
            direction,
            start,
            end,
            normalizedStart,
            normalizedEnd,
            isDragging
        };
    }

    private emitRangeSelection(
        phase: "hover" | "dragging" | "committed" | "cleared",
        start: RangeSelectionAnchor,
        end: RangeSelectionAnchor,
        isDragging: boolean
    ): void {
        if (!this.rangeSelection.callbacks.onChange || !start || !end) {
            return;
        }
        this.rangeSelection.callbacks.onChange(this.createSelectionPayload(phase, start, end, isDragging));
    }

    private renderRangeSelection(): void {
        if (!this.rangeInteractionOverlay) {
            return;
        }
        if (this.needsCommittedRangeAnchorRefresh) {
            this.refreshCommittedRangeAnchorsFromTimestamps();
            this.needsCommittedRangeAnchorRefresh = false;
        }
        const hideSelectionVisuals: boolean = this.shouldHideSelectionRangeVisuals();
        const useMaskGrayOut: boolean = this.shouldUseMaskGrayOut();

        // Mask drag: update pooled rectangles + handle positions in place (no innerHTML wipe per frame).
        if (useMaskGrayOut && this.isRangeDragging && this.dragStartAnchor && this.dragCurrentAnchor) {
            if (this.hasActiveRangeSelectionOpacity) {
                this.resetRangeSelectionNoteOpacity();
            }
            const maskSelection: RangeSelectionPayload = this.createSelectionPayload(
                "dragging",
                this.dragStartAnchor,
                this.dragCurrentAnchor,
                true
            );
            this.applyOutsideMaskSegments(this.buildOutsideMaskSegments(maskSelection));
            if (!hideSelectionVisuals) {
                if (!this.maskDragChromePrepared) {
                    this.clearRangeChromeLayer();
                    this.maskDragChromePrepared = true;
                }
                const lineWidthPx: number = this.getSelectionLineWidthPx();
                const lineColor: string = this.getSelectionLineColor();
                this.updateDragHandleLine(0, this.dragStartAnchor, lineColor, lineWidthPx);
                this.updateDragHandleLine(1, this.dragCurrentAnchor, lineColor, lineWidthPx);
            }
            return;
        }

        this.maskDragChromePrepared = false;
        this.clearRangeChromeLayer();
        if (useMaskGrayOut) {
            if (this.hasActiveRangeSelectionOpacity) {
                this.resetRangeSelectionNoteOpacity();
            }
        } else {
            this.updateRangeSelectionOpacity();
            this.applyOutsideMaskSegments([]);
        }

        if (this.dragStartAnchor && this.dragCurrentAnchor) {
            if (useMaskGrayOut) {
                const maskSelection: RangeSelectionPayload = this.createSelectionPayload(
                    this.isRangeDragging ? "dragging" : "committed",
                    this.dragStartAnchor,
                    this.dragCurrentAnchor,
                    this.isRangeDragging
                );
                this.applyOutsideMaskSegments(this.buildOutsideMaskSegments(maskSelection));
            }
            if (!hideSelectionVisuals && (this.isRangeDragging || this.shouldShowCommittedRangeFill())) {
                this.renderSelectionRangeOverlay(this.dragStartAnchor, this.dragCurrentAnchor);
            }
            if (!hideSelectionVisuals) {
                const lineWidthPx: number = this.getSelectionLineWidthPx();
                this.renderVerticalLine(this.dragStartAnchor, this.getSelectionLineColor(), lineWidthPx);
                this.renderVerticalLine(this.dragCurrentAnchor, this.getSelectionLineColor(), lineWidthPx);
            }
            if (!hideSelectionVisuals && !this.isRangeDragging) {
                const currentSelection: RangeSelectionPayload = this.getRangeSelection();
                this.renderRangeActionButtons(currentSelection);
            }
            return;
        }
        if (!hideSelectionVisuals && !this.isRangeDragging && this.pendingTouchRangeStartAnchor) {
            this.renderVerticalLine(this.pendingTouchRangeStartAnchor, this.getSelectionLineColor(), this.getSelectionLineWidthPx());
            return;
        }
        if (!hideSelectionVisuals && this.shouldShowHoverLine() && !this.isRangeDragging && this.hoverAnchor) {
            this.renderVerticalLine(this.hoverAnchor, this.getSelectionLineColor(), 2);
        }
        if (useMaskGrayOut && !(this.dragStartAnchor && this.dragCurrentAnchor)) {
            this.applyOutsideMaskSegments([]);
        }
    }

    private updateDragHandleLine(
        slot: 0 | 1,
        anchor: RangeSelectionAnchor,
        color: string,
        widthPx: number,
        visualOffsetPx: number = 0
    ): void {
        if (!anchor) {
            return;
        }
        let line: HTMLDivElement = this.dragHandleLines[slot];
        if (!line) {
            line = document.createElement("div");
            line.style.position = "absolute";
            line.style.borderRadius = "999px";
            this.ensureRangeChromeLayer().appendChild(line);
            this.dragHandleLines[slot] = line;
        }
        const overlayHeight: number = this.rangeInteractionOverlay?.clientHeight ?? 0;
        let topPx: number = anchor.yPx;
        let heightPx: number = anchor.heightPx;
        const lineOutsideOverlay: boolean = overlayHeight > 0 && (topPx + heightPx < 0 || topPx > overlayHeight);
        const invalidLineGeometry: boolean = !Number.isFinite(topPx) || !Number.isFinite(heightPx) || heightPx <= 1 || lineOutsideOverlay;
        if (invalidLineGeometry && overlayHeight > 0) {
            topPx = 0;
            heightPx = overlayHeight;
        } else if (overlayHeight > 0) {
            topPx = Math.max(0, Math.min(topPx, overlayHeight - 1));
            heightPx = Math.max(1, Math.min(heightPx, overlayHeight - topPx));
        }
        line.style.left = `${anchor.xPx + visualOffsetPx - widthPx / 2}px`;
        line.style.top = `${topPx}px`;
        line.style.width = `${widthPx}px`;
        line.style.height = `${heightPx}px`;
        line.style.backgroundColor = color;
        line.style.display = "block";
    }

    private refreshCommittedRangeAnchorsFromTimestamps(): void {
        if (this.isRangeDragging || !this.dragStartAnchor || !this.dragCurrentAnchor || !this.graphic) {
            return;
        }
        const hadForwardDirection: boolean = this.dragStartAnchor.timestampReal <= this.dragCurrentAnchor.timestampReal;
        const refreshedStartAnchor: RangeSelectionAnchor = this.createAnchorFromTimestamp(new Fraction(this.dragStartAnchor.timestampReal, 1));
        const refreshedEndAnchor: RangeSelectionAnchor = this.createAnchorFromTimestamp(new Fraction(this.dragCurrentAnchor.timestampReal, 1));
        if (!refreshedStartAnchor || !refreshedEndAnchor) {
            return;
        }
        const refreshedSelection: RangeSelectionPayload = this.applySelectionPadding(
            this.createSelectionPayload("committed", refreshedStartAnchor, refreshedEndAnchor, false),
            "both"
        );
        if (hadForwardDirection) {
            this.dragStartAnchor = refreshedSelection.normalizedStart;
            this.dragCurrentAnchor = refreshedSelection.normalizedEnd;
            return;
        }
        this.dragStartAnchor = refreshedSelection.normalizedEnd;
        this.dragCurrentAnchor = refreshedSelection.normalizedStart;
    }

    private renderSelectionRangeOverlay(start: RangeSelectionAnchor, end: RangeSelectionAnchor): void {
        if (!this.graphic) {
            return;
        }
        const selection: RangeSelectionPayload = this.createSelectionPayload("dragging", start, end, this.isRangeDragging);
        const firstAnchor: RangeSelectionAnchor = selection.normalizedStart;
        const lastAnchor: RangeSelectionAnchor = selection.normalizedEnd;
        const selectedFill: string = this.rangeSelection.options.fillColor ?? "rgba(47, 169, 224, 0.25)";
        const viewport: { topPx: number, bottomPx: number } = this.getRangeSelectionViewportYPx();
        for (const page of this.graphic.MusicPages) {
            for (const system of page.MusicSystems) {
                if (!this.isSystemInRangeSelectionViewport(system, viewport)) {
                    continue;
                }
                const systemIndex: number = this.getSystemIndex(system);
                if (systemIndex < firstAnchor.systemIndex || systemIndex > lastAnchor.systemIndex) {
                    continue;
                }
                const horizontal: { leftPx: number, rightPx: number } = this.getSystemHorizontalBoundsInPixels(system);
                const vertical: { yPx: number, heightPx: number } = this.getSystemVerticalBoundsInPixels(system);
                let selectionLeft: number = horizontal.leftPx;
                let selectionRight: number = horizontal.rightPx;
                if (systemIndex === firstAnchor.systemIndex) {
                    selectionLeft = firstAnchor.xPx;
                }
                if (systemIndex === lastAnchor.systemIndex) {
                    selectionRight = lastAnchor.xPx;
                }
                if (firstAnchor.systemIndex === lastAnchor.systemIndex) {
                    selectionLeft = Math.min(firstAnchor.xPx, lastAnchor.xPx);
                    selectionRight = Math.max(firstAnchor.xPx, lastAnchor.xPx);
                }
                const minLeftPx: number = systemIndex === firstAnchor.systemIndex
                    ? horizontal.leftPx - this.getRangeSelectionPreStartPaddingPx()
                    : horizontal.leftPx;
                selectionLeft = Math.max(minLeftPx, selectionLeft);
                selectionRight = Math.min(horizontal.rightPx, selectionRight);
                if (selectionRight < selectionLeft) {
                    const temp: number = selectionLeft;
                    selectionLeft = selectionRight;
                    selectionRight = temp;
                }
                this.renderRectangle(selectionLeft, vertical.yPx, Math.max(1, selectionRight - selectionLeft), vertical.heightPx, selectedFill);
            }
        }
    }

    private renderVerticalLine(anchor: RangeSelectionAnchor, color: string, widthPx: number, visualOffsetPx: number = 0): void {
        const line: HTMLDivElement = document.createElement("div");
        const overlayHeight: number = this.rangeInteractionOverlay?.clientHeight ?? 0;
        let topPx: number = anchor.yPx;
        let heightPx: number = anchor.heightPx;
        const lineOutsideOverlay: boolean = overlayHeight > 0 && (topPx + heightPx < 0 || topPx > overlayHeight);
        const invalidLineGeometry: boolean = !Number.isFinite(topPx) || !Number.isFinite(heightPx) || heightPx <= 1 || lineOutsideOverlay;
        if (invalidLineGeometry && overlayHeight > 0) {
            // App layouts with additional wrappers/transforms can shift computed Y bounds.
            // Fall back to overlay height so the cursor remains visible.
            topPx = 0;
            heightPx = overlayHeight;
        } else if (overlayHeight > 0) {
            topPx = Math.max(0, Math.min(topPx, overlayHeight - 1));
            heightPx = Math.max(1, Math.min(heightPx, overlayHeight - topPx));
        }
        line.style.position = "absolute";
        line.style.left = `${anchor.xPx + visualOffsetPx - widthPx / 2}px`;
        line.style.top = `${topPx}px`;
        line.style.width = `${widthPx}px`;
        line.style.height = `${heightPx}px`;
        line.style.borderRadius = "999px";
        line.style.backgroundColor = color;
        this.ensureRangeChromeLayer().appendChild(line);
    }

    private renderRectangle(leftPx: number, topPx: number, widthPx: number, heightPx: number, color: string): void {
        if (widthPx <= 0 || heightPx <= 0) {
            return;
        }
        const rect: HTMLDivElement = document.createElement("div");
        rect.style.position = "absolute";
        rect.style.left = `${leftPx}px`;
        rect.style.top = `${topPx}px`;
        rect.style.width = `${widthPx}px`;
        rect.style.height = `${heightPx}px`;
        rect.style.backgroundColor = color;
        this.ensureRangeChromeLayer().appendChild(rect);
    }

    private renderRangeActionButtons(selection: RangeSelectionPayload): void {
        if (!selection || !this.rangeSelection.callbacks.onControlsRender) {
            return;
        }
        const buttonsContainer: HTMLDivElement = document.createElement("div");
        buttonsContainer.style.position = "absolute";
        buttonsContainer.style.pointerEvents = "auto";
        buttonsContainer.style.display = "flex";
        buttonsContainer.style.flexDirection = "column";
        buttonsContainer.style.gap = "6px";
        buttonsContainer.style.zIndex = "9";
        this.rangeSelection.callbacks.onControlsRender(buttonsContainer, selection);
        if (buttonsContainer.childElementCount < 1) {
            return;
        }
        const overlayWidthPx: number = this.rangeInteractionOverlay?.clientWidth ?? 0;
        const overlayHeightPx: number = this.rangeInteractionOverlay?.clientHeight ?? 0;
        const controlsWidthPx: number = buttonsContainer.offsetWidth;
        const controlsHeightPx: number = buttonsContainer.offsetHeight;
        const horizontalMarginPx: number = 10;
        const extraLeftOffsetPx: number = 32;
        const verticalMarginPx: number = 8;

        const startAnchor: RangeSelectionAnchor = selection.normalizedStart;
        const endAnchor: RangeSelectionAnchor = selection.normalizedEnd;
        const sameSystemTolerancePx: number = 1;
        let controlsAnchor: RangeSelectionAnchor = startAnchor;
        if (endAnchor.yPx < startAnchor.yPx - sameSystemTolerancePx) {
            // End handle is on a topmost system.
            controlsAnchor = endAnchor;
        } else if (Math.abs(endAnchor.yPx - startAnchor.yPx) <= sameSystemTolerancePx && endAnchor.xPx < startAnchor.xPx) {
            // Same system: use the leftmost handle.
            controlsAnchor = endAnchor;
        }

        // Place controls to the left of the selected anchor handle.
        const preferredLeftPx: number = controlsAnchor.xPx - controlsWidthPx - horizontalMarginPx - extraLeftOffsetPx;
        const fallbackLeftPx: number = controlsAnchor.xPx + horizontalMarginPx;
        let leftPx: number = preferredLeftPx;
        if (leftPx < horizontalMarginPx) {
            leftPx = fallbackLeftPx;
        }
        if (overlayWidthPx > 0) {
            const maxLeftPx: number = Math.max(horizontalMarginPx, overlayWidthPx - controlsWidthPx - horizontalMarginPx);
            leftPx = Math.min(maxLeftPx, Math.max(horizontalMarginPx, leftPx));
        }

        let topPx: number = Math.max(verticalMarginPx, controlsAnchor.yPx);
        if (overlayHeightPx > 0) {
            const maxTopPx: number = Math.max(verticalMarginPx, overlayHeightPx - controlsHeightPx - verticalMarginPx);
            topPx = Math.min(maxTopPx, Math.max(verticalMarginPx, topPx));
        }

        buttonsContainer.style.left = `${leftPx}px`;
        buttonsContainer.style.top = `${topPx}px`;
        this.ensureRangeChromeLayer().appendChild(buttonsContainer);
    }

    private getSelectionLineColor(): string {
        return this.rangeSelection.options.lineColor ?? "rgba(47, 169, 224, 0.95)";
    }

    private getSelectionLineWidthPx(): number {
        const baseLineWidthPx: number = this.rangeSelection.options.lineWidthPx ?? 12;
        const zoomValue: number = Number.isFinite(this.zoom) ? this.zoom : 1;
        const zoomScaledLineWidthPx: number = baseLineWidthPx * Math.sqrt(Math.max(0.25, zoomValue));
        return Math.max(6, Math.min(14, zoomScaledLineWidthPx));
    }

    private shouldHideSelectionRangeVisuals(): boolean {
        return this.rangeSelection.options.hideSelectionRange === true;
    }

    private shouldShowHoverLine(): boolean {
        return this.rangeSelection.options.showHoverLine !== false;
    }

    private getRangeSelectionViewportYPx(): { topPx: number, bottomPx: number } {
        if (!this.container) {
            return undefined;
        }
        const containerRect: DOMRect = this.container.getBoundingClientRect();
        if (!containerRect || !Number.isFinite(containerRect.top)) {
            return undefined;
        }
        const scrollContainer: HTMLElement | Window = this.rangeViewportScrollTarget ?? this.getRangeSelectionScrollContainer();
        let viewportTopClientPx: number = 0;
        let viewportBottomClientPx: number = window.innerHeight;
        if (!this.isWindowObject(scrollContainer)) {
            const scrollRect: DOMRect = scrollContainer.getBoundingClientRect();
            viewportTopClientPx = scrollRect.top;
            viewportBottomClientPx = scrollRect.bottom;
        }
        const viewportHeightPx: number = Math.max(1, viewportBottomClientPx - viewportTopClientPx);
        const overscanPx: number = Math.max(240, Math.min(900, viewportHeightPx * 0.75));
        return {
            topPx: viewportTopClientPx - containerRect.top - overscanPx,
            bottomPx: viewportBottomClientPx - containerRect.top + overscanPx
        };
    }

    private isSystemInRangeSelectionViewport(system: MusicSystem, viewport: { topPx: number, bottomPx: number }): boolean {
        if (!system || !viewport) {
            return true;
        }
        const vertical: { yPx: number, heightPx: number } = this.getSystemVerticalBoundsInPixels(system);
        return vertical.yPx + vertical.heightPx >= viewport.topPx && vertical.yPx <= viewport.bottomPx;
    }

    private getRangeSelectionVisibleSystemIndexes(
        viewport: { topPx: number, bottomPx: number }
    ): { indexes: Set<number>, indexBySystem: Map<MusicSystem, number> } {
        const indexes: Set<number> = new Set<number>();
        const indexBySystem: Map<MusicSystem, number> = new Map<MusicSystem, number>();
        if (!this.graphic) {
            return { indexes, indexBySystem };
        }
        let systemIndex: number = 0;
        for (const page of this.graphic.MusicPages) {
            for (const system of page.MusicSystems) {
                if (this.isSystemInRangeSelectionViewport(system, viewport)) {
                    indexes.add(systemIndex);
                    indexBySystem.set(system, systemIndex);
                }
                systemIndex++;
            }
        }
        return { indexes, indexBySystem };
    }

    private getSelectionOverlayZIndex(): number {
        return this.rangeSelection.options.overlayZIndex ?? 8;
    }

    private getRangeSelectionPreStartPaddingPx(): number {
        // Allow hovering and dragging slightly before the first visible system timestamp,
        // so notes at the start boundary (timestamp 0) can still be selected.
        return Math.max(this.getSelectionLineWidthPx() * 4, 48);
    }

    private selectionHasAnyNotes(start: RangeSelectionAnchor, end: RangeSelectionAnchor): boolean {
        if (!this.sheet || !start || !end || !this.graphic) {
            return false;
        }
        const selection: RangeSelectionPayload = this.createSelectionPayload("committed", start, end, false);
        const segments: Array<{ systemIndex: number, leftPx: number, rightPx: number }> = this.getSelectionSegments(selection);
        for (const instrument of this.sheet.Instruments) {
            for (const staff of instrument.Staves) {
                for (const voice of staff.Voices) {
                    for (const voiceEntry of voice.VoiceEntries) {
                        for (const note of voiceEntry.Notes) {
                            if (note.isRest()) {
                                continue;
                            }
                            const graphicalNote: GraphicalNote = this.rules.GNote(note);
                            if (!graphicalNote) {
                                continue;
                            }
                            if (this.isGraphicalNoteInSelection(note, graphicalNote, selection, segments)) {
                                return true;
                            }
                        }
                    }
                }
            }
        }
        return false;
    }

    private shouldGrayOutNonSelectedNotes(): boolean {
        return this.rangeSelection.options.grayOutNonSelectedNotes !== false;
    }

    /** When true, dim non-selected areas via cheap overlay rectangles instead of per-note opacity. */
    private shouldUseMaskGrayOut(): boolean {
        if (!this.shouldGrayOutNonSelectedNotes()) {
            return false;
        }
        return this.rangeSelection.options.grayOutStrategy === "mask";
    }

    private getOutsideMaskColor(): string {
        const configuredColor: string = this.rangeSelection.options.outsideMaskColor;
        if (configuredColor) {
            return configuredColor;
        }
        // Transparent dark veil — dims notation behind it without a white wash (closer to per-note opacity).
        const noteOpacity: number = this.getNonSelectedNotesOpacity();
        const dimAlpha: number = Math.min(0.18, Math.max(0.08, (1 - noteOpacity) * 0.14));
        return `rgba(0, 0, 0, ${dimAlpha})`;
    }

    /** Pixels kept fully unmasked past each range handle (outer edge + slop). */
    private getMaskBarClearancePx(edge: "start" | "end"): number {
        const lineWidthPx: number = this.getSelectionLineWidthPx();
        const halfHandlePx: number = lineWidthPx / 2;
        if (edge === "end") {
            return halfHandlePx + 6;
        }
        return halfHandlePx + lineWidthPx / 2 + 8;
    }

    private getMaskClearBoundaryPx(anchor: RangeSelectionAnchor, edge: "start" | "end"): number {
        const visualPx: number = anchor.xPx;
        const selectionPx: number = this.getSelectionBoundaryXPx(anchor);
        const handleOuterPx: number = edge === "end"
            ? Math.max(visualPx, selectionPx ?? visualPx)
            : Math.min(visualPx, selectionPx ?? visualPx);
        const clearancePx: number = this.getMaskBarClearancePx(edge);
        if (edge === "end") {
            // Right handle — keep as-is.
            return handleOuterPx + clearancePx - 8 - 4;
        }
        // Nudge clear boundary right so the left mask meets the start handle without a bright strip.
        return handleOuterPx - clearancePx + 16 + 4;
    }

    private buildOutsideMaskSegments(selection: RangeSelectionPayload): OutsideMaskSegment[] {
        if (!this.graphic || !selection) {
            return [];
        }
        const segments: OutsideMaskSegment[] = [];
        const maskColor: string = this.getOutsideMaskColor();
        const firstSystemIndex: number = selection.normalizedStart.systemIndex;
        const lastSystemIndex: number = selection.normalizedEnd.systemIndex;
        for (const page of this.graphic.MusicPages) {
            for (const system of page.MusicSystems) {
                const systemIndex: number = this.getSystemIndex(system);
                const horizontal: { leftPx: number, rightPx: number } = this.getSystemMaskHorizontalBoundsInPixels(system);
                const vertical: { yPx: number, heightPx: number } = this.getSystemVerticalBoundsInPixels(system);
                const systemWidthPx: number = Math.max(1, horizontal.rightPx - horizontal.leftPx);
                if (systemIndex < firstSystemIndex || systemIndex > lastSystemIndex) {
                    segments.push({
                        leftPx: horizontal.leftPx,
                        topPx: vertical.yPx,
                        widthPx: systemWidthPx,
                        heightPx: vertical.heightPx,
                        color: maskColor
                    });
                    continue;
                }
                let clearLeftPx: number = horizontal.leftPx;
                let clearRightPx: number = horizontal.rightPx;
                if (systemIndex === firstSystemIndex) {
                    clearLeftPx = this.getMaskClearBoundaryPx(selection.normalizedStart, "start");
                }
                if (systemIndex === lastSystemIndex) {
                    clearRightPx = this.getMaskClearBoundaryPx(selection.normalizedEnd, "end");
                }
                if (clearRightPx < clearLeftPx) {
                    const temp: number = clearLeftPx;
                    clearLeftPx = clearRightPx;
                    clearRightPx = temp;
                }
                const maskLeftEdgePx: number = horizontal.leftPx;
                const leftMaskWidthPx: number = clearLeftPx - maskLeftEdgePx;
                if (leftMaskWidthPx > 0.5) {
                    segments.push({
                        leftPx: maskLeftEdgePx,
                        topPx: vertical.yPx,
                        widthPx: leftMaskWidthPx,
                        heightPx: vertical.heightPx,
                        color: maskColor
                    });
                }
                const rightMaskWidthPx: number = horizontal.rightPx - clearRightPx;
                if (rightMaskWidthPx > 0.5) {
                    segments.push({
                        leftPx: clearRightPx,
                        topPx: vertical.yPx,
                        widthPx: rightMaskWidthPx,
                        heightPx: vertical.heightPx,
                        color: maskColor
                    });
                }
            }
        }
        return segments;
    }

    /** Reuses pooled mask divs and only updates geometry — avoids create/destroy churn during handle drag. */
    private applyOutsideMaskSegments(segments: OutsideMaskSegment[]): void {
        if (!this.rangeInteractionOverlay) {
            return;
        }
        const layer: HTMLDivElement = this.ensureOutsideMaskLayer();
        for (let index: number = 0; index < segments.length; index++) {
            const segment: OutsideMaskSegment = segments[index];
            let rect: HTMLDivElement = this.outsideMaskPool[index];
            if (!rect) {
                rect = document.createElement("div");
                rect.className = "osmd-range-outside-mask";
                rect.style.position = "absolute";
                rect.style.pointerEvents = "none";
                // Multiply darkens ink/staff lines in place instead of painting a flat white slab on top.
                rect.style.mixBlendMode = "multiply";
                layer.appendChild(rect);
                this.outsideMaskPool[index] = rect;
            }
            rect.style.display = "block";
            rect.style.left = `${segment.leftPx}px`;
            rect.style.top = `${segment.topPx}px`;
            rect.style.width = `${segment.widthPx}px`;
            rect.style.height = `${segment.heightPx}px`;
            rect.style.backgroundColor = segment.color;
        }
        for (let index: number = segments.length; index < this.outsideMaskPool.length; index++) {
            this.outsideMaskPool[index].style.display = "none";
        }
    }

    private getNonSelectedNotesOpacity(): number {
        return this.rangeSelection.options.nonSelectedNotesOpacity ?? 0.28;
    }

    private getGrayOutUpdateIntervalMs(): number {
        const configuredIntervalMs: number = this.rangeSelection.options.grayOutUpdateIntervalMs ?? 25;
        const baseIntervalMs: number = Math.max(0, configuredIntervalMs);
        if (this.isRangeDragging) {
            // Keep drag interaction smooth by lowering opacity update frequency while dragging.
            // A full-fidelity update is still applied when drag settles/commits.
            return Math.max(80, baseIntervalMs);
        }
        // Committed/settled state: scroll-driven refreshes are throttled by the configured interval so
        // that scrolling a long score stays responsive on low-end devices. A full-fidelity pass still
        // runs once scrolling settles (see scheduleRangeViewportSettleUpdate).
        return baseIntervalMs;
    }

    /** Lightweight scroll/resize refresh for a committed range. Extends the gray-out to systems that
     *  scrolled into view without wiping the overlay or rebuilding action buttons. Throttled by
     *  getGrayOutUpdateIntervalMs(); decorations are deferred until scrolling settles (force=true). */
    private updateRangeSelectionViewport(force: boolean): void {
        if (this.isRangeDragging || !this.dragStartAnchor || !this.dragCurrentAnchor) {
            return;
        }
        if (!this.shouldGrayOutNonSelectedNotes()) {
            this.cancelPendingRangeOpacityUpdate();
            this.resetRangeSelectionNoteOpacity();
            return;
        }
        if (force) {
            this.cancelPendingRangeOpacityUpdate();
            this.applyNoteOpacityForCurrentSelection(true);
            this.lastRangeOpacityUpdateTimestampMs = Date.now();
            return;
        }
        const intervalMs: number = this.getGrayOutUpdateIntervalMs();
        const nowMs: number = Date.now();
        const elapsedMs: number = nowMs - this.lastRangeOpacityUpdateTimestampMs;
        if (intervalMs <= 0 || elapsedMs >= intervalMs) {
            this.cancelPendingRangeOpacityUpdate();
            this.applyNoteOpacityForCurrentSelection(false);
            this.lastRangeOpacityUpdateTimestampMs = nowMs;
            return;
        }
        if (this.rangeOpacityUpdateTimeoutId !== 0) {
            return;
        }
        const remainingMs: number = Math.max(0, intervalMs - elapsedMs);
        this.rangeOpacityUpdateTimeoutId = window.setTimeout((): void => {
            this.rangeOpacityUpdateTimeoutId = 0;
            if (!this.rangeSelection.enabled || this.isRangeDragging || !this.dragStartAnchor || !this.dragCurrentAnchor) {
                return;
            }
            this.applyNoteOpacityForCurrentSelection(false);
            this.lastRangeOpacityUpdateTimestampMs = Date.now();
        }, remainingMs);
    }

    private shouldShowCommittedRangeFill(): boolean {
        return this.rangeSelection.options.showCommittedRangeFill === true;
    }

    private updateRangeSelectionOpacity(): void {
        if (!this.dragStartAnchor || !this.dragCurrentAnchor || !this.shouldGrayOutNonSelectedNotes()) {
            this.cancelPendingRangeOpacityUpdate();
            this.resetRangeSelectionNoteOpacity();
            return;
        }
        const intervalMs: number = this.getGrayOutUpdateIntervalMs();
        if (!this.isRangeDragging || intervalMs <= 0) {
            this.cancelPendingRangeOpacityUpdate();
            this.applyRangeSelectionOpacityNow(true);
            return;
        }
        const nowMs: number = Date.now();
        const elapsedMs: number = nowMs - this.lastRangeOpacityUpdateTimestampMs;
        if (!this.hasActiveRangeSelectionOpacity || elapsedMs >= intervalMs) {
            this.cancelPendingRangeOpacityUpdate();
            this.applyRangeSelectionOpacityNow();
            return;
        }
        if (this.rangeOpacityUpdateTimeoutId !== 0) {
            return;
        }
        const remainingMs: number = Math.max(0, intervalMs - elapsedMs);
        this.rangeOpacityUpdateTimeoutId = window.setTimeout((): void => {
            this.rangeOpacityUpdateTimeoutId = 0;
            if (!this.rangeSelection.enabled || !this.dragStartAnchor || !this.dragCurrentAnchor) {
                this.resetRangeSelectionNoteOpacity();
                return;
            }
            this.applyRangeSelectionOpacityNow();
        }, remainingMs);
    }

    private applyRangeSelectionOpacityNow(forceDecorationUpdate: boolean = false): void {
        const nowMs: number = Date.now();
        const shouldUpdateDecorations: boolean = this.shouldUpdateDecorationOpacity(nowMs, forceDecorationUpdate);
        this.applyNoteOpacityForCurrentSelection(shouldUpdateDecorations);
        this.lastRangeOpacityUpdateTimestampMs = nowMs;
    }

    private shouldUpdateDecorationOpacity(nowMs: number, forceDecorationUpdate: boolean): boolean {
        if (forceDecorationUpdate || !this.isRangeDragging || !this.hasActiveRangeSelectionOpacity) {
            return true;
        }
        // Decorations (expressions, tuplets, structural SVG) are expensive to mutate repeatedly.
        // Defer them during drag and apply at the end via forceDecorationUpdate path.
        return false;
    }

    private cancelPendingRangeOpacityUpdate(): void {
        if (this.rangeOpacityUpdateTimeoutId !== 0) {
            window.clearTimeout(this.rangeOpacityUpdateTimeoutId);
            this.rangeOpacityUpdateTimeoutId = 0;
        }
    }

    private applyNoteOpacityForCurrentSelection(includeDecorations: boolean = true): void {
        if (!this.sheet || !this.graphic || !this.shouldGrayOutNonSelectedNotes()) {
            this.resetRangeSelectionNoteOpacity();
            return;
        }
        if (!this.dragStartAnchor || !this.dragCurrentAnchor) {
            this.resetRangeSelectionNoteOpacity();
            return;
        }
        const selection: RangeSelectionPayload = this.createSelectionPayload("committed", this.dragStartAnchor, this.dragCurrentAnchor, this.isRangeDragging);
        const viewport: { topPx: number, bottomPx: number } = this.getRangeSelectionViewportYPx();
        const visibleSystems: { indexes: Set<number>, indexBySystem: Map<MusicSystem, number> } =
            this.getRangeSelectionVisibleSystemIndexes(viewport);
        // Skip the (expensive) recompute when nothing relevant changed since the last applied gray-out.
        // This makes scroll cheap: while scrolling within the same set of visible systems the key is
        // unchanged, so we bail out before iterating notes or querying decoration elements. The
        // decoration flag is tracked separately so a note-only scroll pass never downgrades a viewport
        // that already had its decorations grayed.
        const nonSelectedOpacity: number = this.getNonSelectedNotesOpacity();
        const visibleSystemsKey: string = Array.from(visibleSystems.indexes).sort((a: number, b: number): number => a - b).join(",");
        const opacityKey: string =
            `${selection.normalizedStart.timestampReal}_${selection.normalizedEnd.timestampReal}`
            + `|${visibleSystemsKey}|${nonSelectedOpacity}`;
        if (
            this.hasActiveRangeSelectionOpacity
            && opacityKey === this.lastRangeOpacityKey
            && (!includeDecorations || this.lastRangeOpacityDecorationsApplied)
        ) {
            return;
        }
        // Always compute full (un-truncated) segments. Viewport culling is applied at the per-element
        // level below; segments must remain complete so that selectionEndBoundary used in
        // applyStructuralElementOpacityForSelection still reflects the true end of the selection
        // (otherwise end-of-line clefs/braces past the visible viewport would be incorrectly grayed).
        const segments: Array<{ systemIndex: number, leftPx: number, rightPx: number }> =
            this.getSelectionSegments(selection);
        const opacityByTarget: Map<string, { graphicalNote: GraphicalNote, shouldHighlight: boolean }> =
            new Map<string, { graphicalNote: GraphicalNote, shouldHighlight: boolean }>();
        const noteheadOpacityByTarget: Map<string, { graphicalNote: GraphicalNote, shouldHighlight: boolean }> =
            new Map<string, { graphicalNote: GraphicalNote, shouldHighlight: boolean }>();

        for (const verticalContainer of this.graphic.VerticalGraphicalStaffEntryContainers) {
            for (const staffEntry of verticalContainer?.StaffEntries ?? []) {
                if (!staffEntry) {
                    continue;
                }
                const system: MusicSystem = staffEntry.parentMeasure?.ParentMusicSystem;
                const systemIndex: number | undefined = visibleSystems.indexBySystem.get(system);
                if (systemIndex === undefined) {
                    continue;
                }
                for (const graphicalVoiceEntry of staffEntry.graphicalVoiceEntries ?? []) {
                    for (const graphicalNote of graphicalVoiceEntry?.notes ?? []) {
                        const note: any = graphicalNote?.sourceNote;
                        if (!graphicalNote || !note) {
                            continue;
                        }
                        const noteIsSelected: boolean = this.isGraphicalNoteInSelection(
                            note,
                            graphicalNote,
                            selection,
                            segments,
                            systemIndex
                        );
                        const isTieStopNote: boolean = this.isRangeDragging
                            ? false
                            : this.isTieStopNote(note, selection, segments);
                        const shouldHighlight: boolean = noteIsSelected && !isTieStopNote;
                        const targetKey: string = this.getGraphicalOpacityTargetKey(graphicalNote, note);
                        const existingTargetState: { graphicalNote: GraphicalNote, shouldHighlight: boolean } =
                            opacityByTarget.get(targetKey);
                        if (existingTargetState) {
                            existingTargetState.shouldHighlight = existingTargetState.shouldHighlight || shouldHighlight;
                        } else {
                            opacityByTarget.set(targetKey, { graphicalNote, shouldHighlight });
                        }
                        const noteheadKey: string = this.getGraphicalNoteheadOpacityTargetKey(graphicalNote, note);
                        noteheadOpacityByTarget.set(noteheadKey, { graphicalNote, shouldHighlight });
                    }
                }
            }
        }
        for (const targetState of opacityByTarget.values()) {
            targetState.graphicalNote.setOpacity(targetState.shouldHighlight ? 1.0 : nonSelectedOpacity);
            this.rangeOpacityTouchedGraphicalNotes.add(targetState.graphicalNote);
        }
        const shouldUpdateSpecificNoteheads: boolean = !this.isRangeDragging;
        if (shouldUpdateSpecificNoteheads) {
            for (const noteheadState of noteheadOpacityByTarget.values()) {
                this.setSpecificGraphicalNoteheadOpacity(
                    noteheadState.graphicalNote,
                    noteheadState.shouldHighlight ? 1.0 : nonSelectedOpacity
                );
            }
        }
        if (includeDecorations) {
            this.applyStaffEntryElementOpacityForSelection(segments, nonSelectedOpacity, visibleSystems.indexBySystem);
            this.applyMeasureNumberOpacityForSelection(segments, nonSelectedOpacity, visibleSystems.indexes);
            this.applyExpressionOpacityForSelection(segments, nonSelectedOpacity, visibleSystems.indexes);
            this.applyStructuralElementOpacityForSelection(segments, nonSelectedOpacity, viewport);
            this.applyTupletOpacityForSelection(segments, nonSelectedOpacity, viewport);
        }
        this.hasActiveRangeSelectionOpacity = true;
        this.lastRangeOpacityKey = opacityKey;
        this.lastRangeOpacityDecorationsApplied = includeDecorations;
    }

    private resetRangeSelectionNoteOpacity(): void {
        this.lastRangeOpacityKey = "";
        this.lastRangeOpacityDecorationsApplied = false;
        if (!this.hasActiveRangeSelectionOpacity) {
            return;
        }
        if (!this.sheet || !this.graphic) {
            this.rangeOpacityTouchedGraphicalNotes.clear();
            this.rangeOpacityTouchedNoteheadElements.clear();
            this.rangeOpacityTouchedElements.clear();
            this.hasActiveRangeSelectionOpacity = false;
            return;
        }
        for (const graphicalNote of this.rangeOpacityTouchedGraphicalNotes) {
            graphicalNote.setOpacity(1.0);
        }
        for (const noteheadElement of this.rangeOpacityTouchedNoteheadElements) {
            if (noteheadElement?.isConnected) {
                noteheadElement.setAttribute("opacity", "1");
            }
        }
        for (const element of this.rangeOpacityTouchedElements) {
            if (element?.isConnected) {
                element.setAttribute("opacity", "1");
            }
        }
        this.rangeOpacityTouchedGraphicalNotes.clear();
        this.rangeOpacityTouchedNoteheadElements.clear();
        this.rangeOpacityTouchedElements.clear();
        this.hasActiveRangeSelectionOpacity = false;
        this.lastRangeOpacityUpdateTimestampMs = 0;
    }

    private applyStaffEntryElementOpacityForSelection(
        segments: Array<{ systemIndex: number, leftPx: number, rightPx: number }>,
        nonSelectedOpacity: number,
        visibleSystemIndexBySystem: Map<MusicSystem, number>
    ): void {
        if (!this.graphic) {
            return;
        }
        for (const verticalContainer of this.graphic.VerticalGraphicalStaffEntryContainers) {
            for (const staffEntry of verticalContainer?.StaffEntries ?? []) {
                if (!staffEntry) {
                    continue;
                }
                const systemIndex: number | undefined = visibleSystemIndexBySystem.get(staffEntry.parentMeasure?.ParentMusicSystem);
                if (systemIndex === undefined) {
                    continue;
                }
                const entryXPx: number = staffEntry.PositionAndShape.AbsolutePosition.x * this.zoom * 10.0;
                const opacity: number = this.isXInSelection(systemIndex, entryXPx, segments) ? 1.0 : nonSelectedOpacity;
                for (const lyricsEntry of staffEntry.LyricsEntries ?? []) {
                    this.setGraphicalLabelOpacity(lyricsEntry?.GraphicalLabel, opacity);
                }
                for (const fingeringEntry of staffEntry.FingeringEntries ?? []) {
                    this.setGraphicalLabelOpacity(fingeringEntry, opacity);
                }
                for (const chordContainer of staffEntry.graphicalChordContainers ?? []) {
                    this.setGraphicalLabelOpacity(chordContainer?.GraphicalLabel, opacity);
                }
            }
        }
    }

    private setGraphicalLabelOpacity(label: GraphicalLabel, opacity: number): void {
        if (!label?.SVGNode) {
            return;
        }
        const labelNode: Element = label.SVGNode as Element;
        this.setRangeSelectionElementOpacity(labelNode, opacity);
    }

    private setRangeSelectionElementOpacity(element: Element, opacity: number): void {
        if (!element) {
            return;
        }
        element.setAttribute("opacity", opacity.toString());
        this.rangeOpacityTouchedElements.add(element);
    }

    private applyMeasureNumberOpacityForSelection(
        segments: Array<{ systemIndex: number, leftPx: number, rightPx: number }>,
        nonSelectedOpacity: number,
        visibleSystemIndexes: Set<number>
    ): void {
        if (!this.graphic) {
            return;
        }
        const nonNoteSelectionTolerancePx: number = this.getNonNoteSelectionTolerancePx();
        for (const page of this.graphic.MusicPages) {
            for (const system of page.MusicSystems) {
                const systemIndex: number = this.getSystemIndex(system);
                if (!visibleSystemIndexes.has(systemIndex)) {
                    continue;
                }
                for (const measureNumberLabel of system.MeasureNumberLabels ?? []) {
                    const centerXPx: number = this.getGraphicalObjectCenterXPx(measureNumberLabel as any);
                    const opacity: number = this.isXInSelection(
                        systemIndex,
                        centerXPx,
                        segments,
                        nonNoteSelectionTolerancePx
                    ) ? 1.0 : nonSelectedOpacity;
                    this.setGraphicalLabelOpacity(measureNumberLabel, opacity);
                }
            }
        }
    }

    private applyExpressionOpacityForSelection(
        segments: Array<{ systemIndex: number, leftPx: number, rightPx: number }>,
        nonSelectedOpacity: number,
        visibleSystemIndexes: Set<number>
    ): void {
        if (!this.graphic) {
            return;
        }
        const nonNoteSelectionTolerancePx: number = this.getNonNoteSelectionTolerancePx();
        for (const page of this.graphic.MusicPages) {
            for (const system of page.MusicSystems) {
                const systemIndex: number = this.getSystemIndex(system);
                if (!visibleSystemIndexes.has(systemIndex)) {
                    continue;
                }
                for (const staffLine of system.StaffLines ?? []) {
                    for (const expression of staffLine.AbstractExpressions ?? []) {
                        const expressionXPx: number = this.getExpressionCenterXPx(expression);
                        const opacity: number = this.isXInSelection(
                            systemIndex,
                            expressionXPx,
                            segments,
                            nonNoteSelectionTolerancePx
                        ) ? 1.0 : nonSelectedOpacity;
                        this.setGraphicalLabelOpacity(expression?.Label, opacity);
                        this.setExpressionOpacity(expression, opacity);
                    }
                }
            }
        }
    }

    private setExpressionOpacity(expression: AbstractGraphicalExpression, opacity: number): void {
        if (!expression) {
            return;
        }
        const expressionNode: Element = (expression?.PositionAndShape as any)?.SVGNode as Element;
        if (expressionNode) {
            this.setRangeSelectionElementOpacity(expressionNode, opacity);
        }
        const expressionLines: any[] = (expression as any)?.Lines ?? [];
        for (const line of expressionLines) {
            const lineElement: Element = line?.SVGElement as Element;
            if (lineElement) {
                this.setRangeSelectionElementOpacity(lineElement, opacity);
            }
        }
    }

    private getExpressionCenterXPx(expression: AbstractGraphicalExpression): number {
        const expressionLabel: GraphicalLabel = expression?.Label;
        const labelCenterXPx: number = this.getGraphicalObjectCenterXPx(expressionLabel as any);
        if (Number.isFinite(labelCenterXPx)) {
            return labelCenterXPx;
        }
        const expressionCenterXPx: number = this.getGraphicalObjectCenterXPx(expression as any);
        if (Number.isFinite(expressionCenterXPx)) {
            return expressionCenterXPx;
        }
        const expressionLines: any[] = (expression as any)?.Lines ?? [];
        for (const line of expressionLines) {
            const startX: number = line?.Start?.x;
            const endX: number = line?.End?.x;
            if (!Number.isFinite(startX) || !Number.isFinite(endX)) {
                continue;
            }
            return ((startX + endX) / 2) * this.zoom * 10.0;
        }
        return Number.NEGATIVE_INFINITY;
    }

    private getGraphicalObjectCenterXPx(graphicalObject: { PositionAndShape?: any }): number {
        const positionAndShape: any = graphicalObject?.PositionAndShape;
        const absoluteX: number = positionAndShape?.AbsolutePosition?.x;
        if (!Number.isFinite(absoluteX)) {
            return Number.NaN;
        }
        const borderLeft: number = positionAndShape?.BorderLeft ?? 0;
        const borderRight: number = positionAndShape?.BorderRight ?? 0;
        return (absoluteX + (borderLeft + borderRight) / 2) * this.zoom * 10.0;
    }

    private applyStructuralElementOpacityForSelection(
        segments: Array<{ systemIndex: number, leftPx: number, rightPx: number }>,
        nonSelectedOpacity: number,
        viewport: { topPx: number, bottomPx: number }
    ): void {
        const elements: SVGGraphicsElement[] = this.getStructuralSelectionElements();
        const containerRect: DOMRect = this.container.getBoundingClientRect();
        const nonNoteSelectionTolerancePx: number = this.getNonNoteSelectionTolerancePx();
        const selectionEndBoundary: { systemIndex: number, xPx: number } = this.getSelectionEndBoundary(segments);
        const clefBraceRects: DOMRect[] = [];
        const elementMetadata: Array<{
            element: SVGGraphicsElement;
            elementRect: DOMRect;
            centerXPx: number;
            system: MusicSystem;
            systemIndex: number;
        }> = [];
        const connectorExtremesBySystem: Map<number, { minXPx: number, maxXPx: number }> = new Map();
        for (const element of elements) {
            const elementRect: DOMRect = element.getBoundingClientRect();
            if (!this.isElementRectInRangeSelectionViewport(elementRect, containerRect, viewport)) {
                continue;
            }
            if (this.isRightSideOnlyStructuralElement(element)) {
                clefBraceRects.push(elementRect);
            }
            const centerXPx: number = (elementRect.left - containerRect.left) + (elementRect.width / 2);
            const centerYPx: number = (elementRect.top - containerRect.top) + (elementRect.height / 2);
            const system: MusicSystem = this.findSystemAtPosition(new PointF2D(centerXPx / (this.zoom * 10.0), centerYPx / (this.zoom * 10.0)));
            const systemIndex: number = this.getSystemIndex(system);
            elementMetadata.push({ element, elementRect, centerXPx, system, systemIndex });
            if (!this.isConnectorStructuralElement(element) || systemIndex < 0 || !Number.isFinite(centerXPx)) {
                continue;
            }
            const extremes: { minXPx: number, maxXPx: number } = connectorExtremesBySystem.get(systemIndex)
                ?? { minXPx: centerXPx, maxXPx: centerXPx };
            extremes.minXPx = Math.min(extremes.minXPx, centerXPx);
            extremes.maxXPx = Math.max(extremes.maxXPx, centerXPx);
            connectorExtremesBySystem.set(systemIndex, extremes);
        }
        for (const meta of elementMetadata) {
            const { element, elementRect, centerXPx, system, systemIndex } = meta;
            const keepBoundaryConnectorVisible: boolean = this.shouldKeepBoundaryConnectorVisible(
                element,
                system,
                centerXPx,
                nonNoteSelectionTolerancePx,
                systemIndex,
                connectorExtremesBySystem
            );
            const isRightSideOnlyElement: boolean = this.isRightSideOnlyStructuralElement(element)
                || (this.isConnectorStructuralElement(element)
                    && this.isConnectorNearClefOrBrace(elementRect, clefBraceRects));
            const opacity: number = keepBoundaryConnectorVisible
                ? 0.6
                : (isRightSideOnlyElement
                    ? (this.isElementAfterSelectionEnd(systemIndex, centerXPx, selectionEndBoundary, nonNoteSelectionTolerancePx)
                        ? nonSelectedOpacity
                        : 1.0)
                    : (this.isXInSelection(
                        systemIndex,
                        centerXPx,
                        segments,
                        nonNoteSelectionTolerancePx
                    ) ? 1.0 : nonSelectedOpacity));
            this.setRangeSelectionElementOpacity(element, opacity);
        }
    }

    private isElementRectInRangeSelectionViewport(
        elementRect: DOMRect,
        containerRect: DOMRect,
        viewport: { topPx: number, bottomPx: number }
    ): boolean {
        if (!elementRect || !containerRect || !viewport) {
            return true;
        }
        const topPx: number = elementRect.top - containerRect.top;
        const bottomPx: number = elementRect.bottom - containerRect.top;
        return bottomPx >= viewport.topPx && topPx <= viewport.bottomPx;
    }

    private isRightSideOnlyStructuralElement(element: SVGGraphicsElement): boolean {
        if (!element) {
            return false;
        }
        if (element.matches(
            ".vf-clef, .vf-stave-clef, [class*='clef'], [id*='clef'], " +
            ".vf-brace, [class*='brace'], [id*='brace']"
        )) {
            return true;
        }
        return Boolean(
            element.closest(
                ".vf-clef, .vf-stave-clef, [class*='clef'], [id*='clef'], " +
                ".vf-brace, [class*='brace'], [id*='brace']"
            )
        );
    }

    private isConnectorStructuralElement(element: SVGGraphicsElement): boolean {
        if (!element) {
            return false;
        }
        if (element.matches(".vf-connector, [class*='connector'], [id*='connector']")) {
            return true;
        }
        return Boolean(element.closest(".vf-connector, [class*='connector'], [id*='connector']"));
    }

    private isConnectorNearClefOrBrace(connectorRect: DOMRect, clefBraceRects: DOMRect[]): boolean {
        if (!connectorRect || !clefBraceRects?.length) {
            return false;
        }
        const connectorCenterX: number = connectorRect.left + (connectorRect.width / 2);
        for (const anchorRect of clefBraceRects) {
            const anchorCenterX: number = anchorRect.left + (anchorRect.width / 2);
            const horizontalDistancePx: number = Math.abs(connectorCenterX - anchorCenterX);
            if (horizontalDistancePx > 72) {
                continue;
            }
            const verticalOverlapPx: number =
                Math.min(connectorRect.bottom, anchorRect.bottom) - Math.max(connectorRect.top, anchorRect.top);
            if (verticalOverlapPx >= -36) {
                return true;
            }
        }
        return false;
    }

    private shouldKeepBoundaryConnectorVisible(
        element: SVGGraphicsElement,
        system: MusicSystem,
        centerXPx: number,
        tolerancePx: number,
        systemIndex: number,
        connectorExtremesBySystem: Map<number, { minXPx: number, maxXPx: number }>
    ): boolean {
        if (this.rangeSelection.options?.keepBoundaryConnectorsVisible === false) {
            return false;
        }
        if (!this.isConnectorStructuralElement(element) || !system || !Number.isFinite(centerXPx)) {
            return false;
        }
        const connectorExtremes: { minXPx: number, maxXPx: number } = connectorExtremesBySystem.get(systemIndex);
        if (!connectorExtremes) {
            return false;
        }
        const boundaryTolerancePx: number = Math.max(16, tolerancePx * 2);
        const nearLeftmostConnector: boolean = Math.abs(centerXPx - connectorExtremes.minXPx) <= boundaryTolerancePx;
        const nearRightmostConnector: boolean = Math.abs(centerXPx - connectorExtremes.maxXPx) <= boundaryTolerancePx;
        return nearLeftmostConnector || nearRightmostConnector;
    }

    private getSelectionEndBoundary(
        segments: Array<{ systemIndex: number, leftPx: number, rightPx: number }>
    ): { systemIndex: number, xPx: number } {
        let endSystemIndex: number = Number.NEGATIVE_INFINITY;
        let endXPx: number = Number.NEGATIVE_INFINITY;
        for (const segment of segments) {
            if (segment.systemIndex > endSystemIndex) {
                endSystemIndex = segment.systemIndex;
                endXPx = segment.rightPx;
                continue;
            }
            if (segment.systemIndex === endSystemIndex) {
                endXPx = Math.max(endXPx, segment.rightPx);
            }
        }
        return { systemIndex: endSystemIndex, xPx: endXPx };
    }

    private isElementAfterSelectionEnd(
        elementSystemIndex: number,
        elementXPx: number,
        selectionEndBoundary: { systemIndex: number, xPx: number },
        tolerancePx: number
    ): boolean {
        if (!selectionEndBoundary || !Number.isFinite(selectionEndBoundary.systemIndex)) {
            return false;
        }
        if (elementSystemIndex > selectionEndBoundary.systemIndex) {
            return true;
        }
        if (elementSystemIndex < selectionEndBoundary.systemIndex) {
            return false;
        }
        return elementXPx > selectionEndBoundary.xPx + tolerancePx;
    }

    private getSystemIndexForGraphicalNote(graphicalNote: GraphicalNote): number {
        const system: MusicSystem = graphicalNote?.parentVoiceEntry?.parentStaffEntry?.parentMeasure?.ParentMusicSystem;
        return this.getSystemIndex(system);
    }

    private getSelectionSegments(
        selection: RangeSelectionPayload
    ): Array<{ systemIndex: number, leftPx: number, rightPx: number }> {
        const segments: Array<{ systemIndex: number, leftPx: number, rightPx: number }> = [];
        if (!this.graphic || !selection) {
            return segments;
        }
        for (const page of this.graphic.MusicPages) {
            for (const system of page.MusicSystems) {
                const systemIndex: number = this.getSystemIndex(system);
                if (systemIndex < selection.normalizedStart.systemIndex || systemIndex > selection.normalizedEnd.systemIndex) {
                    continue;
                }
                const horizontal: { leftPx: number, rightPx: number } = this.getSystemHorizontalBoundsInPixels(system);
                let leftPx: number = horizontal.leftPx;
                let rightPx: number = horizontal.rightPx;
                if (systemIndex === selection.normalizedStart.systemIndex) {
                    leftPx = this.getSelectionBoundaryXPx(selection.normalizedStart);
                }
                if (systemIndex === selection.normalizedEnd.systemIndex) {
                    rightPx = this.getSelectionBoundaryXPx(selection.normalizedEnd);
                }
                const minLeftPx: number = systemIndex === selection.normalizedStart.systemIndex
                    ? horizontal.leftPx - this.getRangeSelectionPreStartPaddingPx()
                    : horizontal.leftPx;
                leftPx = Math.max(minLeftPx, leftPx);
                rightPx = Math.min(horizontal.rightPx, rightPx);
                if (rightPx < leftPx) {
                    const temp: number = leftPx;
                    leftPx = rightPx;
                    rightPx = temp;
                }
                segments.push({ systemIndex, leftPx, rightPx });
            }
        }
        return segments;
    }

    private getSelectionBoundaryXPx(anchor: RangeSelectionAnchor): number {
        if (!anchor) {
            return undefined;
        }
        return Number.isFinite(anchor.selectionXPx) ? anchor.selectionXPx : anchor.xPx;
    }

    private isXInSelection(
        systemIndex: number,
        xPx: number,
        segments: Array<{ systemIndex: number, leftPx: number, rightPx: number }>,
        tolerancePx: number = 0
    ): boolean {
        if (systemIndex < 0) {
            return false;
        }
        for (const segment of segments) {
            if (segment.systemIndex !== systemIndex) {
                continue;
            }
            if (xPx >= segment.leftPx - tolerancePx && xPx <= segment.rightPx + tolerancePx) {
                return true;
            }
        }
        return false;
    }

    private getNonNoteSelectionTolerancePx(): number {
        if (this.rangeSelection.options?.snapToNotes) {
            // Reduce over-dimming around snapped note boundaries for structural symbols.
            return 14;
        }
        return 6;
    }

    private isGraphicalNoteInSelection(
        note: any,
        graphicalNote: GraphicalNote,
        selection: RangeSelectionPayload,
        segments: Array<{ systemIndex: number, leftPx: number, rightPx: number }>,
        knownSystemIndex?: number
    ): boolean {
        const noteSystemIndex: number = knownSystemIndex ?? this.getSystemIndexForGraphicalNote(graphicalNote);
        const noteXPx: number = this.getGraphicalNoteSelectionXPx(graphicalNote);
        if (this.isXInSelection(noteSystemIndex, noteXPx, segments)) {
            return true;
        }
        // When snapToNotes is on, xPx is the authoritative boundary.
        // Don't fall back to timestamp — it re-includes notes at measure boundaries
        // that xPx correctly excluded.
        if (this.rangeSelection.options.snapToNotes) {
            return false;
        }
        const noteTimestampReal: number = this.getAbsoluteTimestampRealForNote(note);
        return this.isTimestampRealInSelection(noteTimestampReal, selection);
    }

    private getGraphicalNoteSelectionXPx(graphicalNote: GraphicalNote): number {
        const zoomScale: number = this.zoom * 10.0;
        const noteXPx: number = graphicalNote?.PositionAndShape?.AbsolutePosition?.x * zoomScale;
        if (!this.rangeSelection.options?.snapToNotes) {
            return noteXPx;
        }
        // Snap boundaries are computed from staff-entry semantic x positions.
        // Use the same semantic x for note inclusion so gray-out matches snap behavior.
        const entryX: number = graphicalNote?.parentVoiceEntry?.parentStaffEntry?.PositionAndShape?.AbsolutePosition?.x;
        const entryXPx: number = entryX * zoomScale;
        return Number.isFinite(entryXPx) ? entryXPx : noteXPx;
    }

    private isTieStopNote(
        note: any,
        selection?: RangeSelectionPayload,
        segments?: Array<{ systemIndex: number, leftPx: number, rightPx: number }>
    ): boolean {
        const tieNotes: any[] = note?.NoteTie?.Notes;
        if (!tieNotes || tieNotes.length < 2) {
            return false;
        }
        const noteIndex: number = tieNotes.indexOf(note);
        if (noteIndex <= 0) {
            return false;
        }
        if (!selection || !segments) {
            return true;
        }
        const parentTieNote: any = tieNotes[noteIndex - 1];
        if (!parentTieNote) {
            return true;
        }
        const parentTieGraphicalNote: GraphicalNote = this.rules.GNote(parentTieNote);
        if (parentTieGraphicalNote) {
            const parentTieInSelection: boolean = this.isGraphicalNoteInSelection(
                parentTieNote,
                parentTieGraphicalNote,
                selection,
                segments
            );
            return !parentTieInSelection;
        }
        const parentTieTimestampReal: number = this.getAbsoluteTimestampRealForNote(parentTieNote);
        const parentTieInSelectionByTimestamp: boolean = this.isTimestampRealInSelection(parentTieTimestampReal, selection);
        return !parentTieInSelectionByTimestamp;
    }

    private getGraphicalOpacityTargetKey(graphicalNote: GraphicalNote, note: any): string {
        const svgId: string = (graphicalNote as any)?.getSVGId?.();
        if (svgId) {
            return `svg:${svgId}`;
        }
        const noteObjectId: number = graphicalNote?.sourceNote?.NoteToGraphicalNoteObjectId;
        if (Number.isFinite(noteObjectId)) {
            return `note:${noteObjectId}`;
        }
        const fallbackTimestampReal: number = this.getAbsoluteTimestampRealForNote(note);
        return `fallback:${fallbackTimestampReal}:${(graphicalNote as any)?.vfnoteIndex ?? -1}`;
    }

    private getGraphicalNoteheadOpacityTargetKey(graphicalNote: GraphicalNote, note: any): string {
        const noteheadIndex: number = (graphicalNote as any)?.vfnoteIndex;
        const targetKey: string = this.getGraphicalOpacityTargetKey(graphicalNote, note);
        return `notehead:${targetKey}:${Number.isFinite(noteheadIndex) ? noteheadIndex : -1}`;
    }

    private setSpecificGraphicalNoteheadOpacity(graphicalNote: GraphicalNote, opacity: number): void {
        const noteheadIndex: number = (graphicalNote as any)?.vfnoteIndex;
        if (!Number.isFinite(noteheadIndex) || noteheadIndex < 0) {
            return;
        }
        const noteheadSvgs: Element[] = (graphicalNote as any)?.getNoteheadSVGs?.() ?? [];
        const noteheadSvg: Element = noteheadSvgs[noteheadIndex];
        if (!noteheadSvg) {
            return;
        }
        const noteheadPaths: NodeListOf<SVGElement> = noteheadSvg.querySelectorAll<SVGElement>("path");
        for (const noteheadPath of noteheadPaths) {
            noteheadPath.setAttribute("opacity", opacity.toString());
            this.rangeOpacityTouchedNoteheadElements.add(noteheadPath);
        }
    }

    private getAbsoluteTimestampRealForNote(note: any): number {
        const timestamp: Fraction = note?.ParentVoiceEntry?.Timestamp;
        const sourceMeasureAbsoluteTimestamp: Fraction =
            note?.ParentVoiceEntry?.ParentSourceStaffEntry?.VerticalContainerParent?.ParentMeasure?.AbsoluteTimestamp
            ?? note?.SourceMeasure?.AbsoluteTimestamp;
        if (!timestamp || !sourceMeasureAbsoluteTimestamp) {
            return undefined;
        }
        return Fraction.plus(sourceMeasureAbsoluteTimestamp, timestamp).RealValue;
    }

    private isTimestampRealInSelection(timestampReal: number, selection: RangeSelectionPayload): boolean {
        if (timestampReal === undefined || !selection?.normalizedStart || !selection?.normalizedEnd) {
            return false;
        }
        const epsilon: number = Fraction.FloatInaccuracyTolerance;
        return timestampReal >= selection.normalizedStart.timestampReal - epsilon
            && timestampReal <= selection.normalizedEnd.timestampReal + epsilon;
    }

    private applyTupletOpacityForSelection(
        segments: Array<{ systemIndex: number, leftPx: number, rightPx: number }>,
        nonSelectedOpacity: number,
        viewport: { topPx: number, bottomPx: number }
    ): void {
        const tupletElements: SVGGraphicsElement[] = this.getTupletElements();
        const containerRect: DOMRect = this.container.getBoundingClientRect();
        const nonNoteSelectionTolerancePx: number = this.getNonNoteSelectionTolerancePx();
        for (const element of tupletElements) {
            const elementRect: DOMRect = element.getBoundingClientRect();
            if (!this.isElementRectInRangeSelectionViewport(elementRect, containerRect, viewport)) {
                continue;
            }
            const centerXPx: number = (elementRect.left - containerRect.left) + (elementRect.width / 2);
            const centerYPx: number = (elementRect.top - containerRect.top) + (elementRect.height / 2);
            const system: MusicSystem = this.findSystemAtPosition(new PointF2D(centerXPx / (this.zoom * 10.0), centerYPx / (this.zoom * 10.0)));
            const systemIndex: number = this.getSystemIndex(system);
            const opacity: number = this.isXInSelection(
                systemIndex,
                centerXPx,
                segments,
                nonNoteSelectionTolerancePx
            ) ? 1.0 : nonSelectedOpacity;
            this.setRangeSelectionElementOpacity(element, opacity);
        }
    }

    private getTupletElements(): SVGGraphicsElement[] {
        return this.rangeSelectionElementCollector.getTupletElements(this.drawer);
    }

    private getStructuralSelectionElements(): SVGGraphicsElement[] {
        return this.rangeSelectionElementCollector.getStructuralElements(this.drawer);
    }

    private setStaffOpacity(staffIndex: number, opacity: number, fallbackOpacity: number): void {
        if (!this.isValidStaffIndex(staffIndex)) {
            return;
        }
        const resolvedOpacity: number = this.clampOpacity(opacity, fallbackOpacity);
        if (resolvedOpacity >= 1) {
            this.staffOpacityOverrides.delete(staffIndex);
        } else {
            this.staffOpacityOverrides.set(staffIndex, resolvedOpacity);
        }
        this.applyOpacityToStaffElements(staffIndex, resolvedOpacity);
    }

    private applyOpacityToStaffElements(staffIndex: number, opacity: number, withinRoots?: ParentNode[]): void {
        const stafflineElements: SVGGElement[] = this.getStafflineElements(staffIndex, withinRoots);
        if (stafflineElements.length === 0) {
            return;
        }
        const opacityString: string = opacity.toString();
        for (const stafflineElement of stafflineElements) {
            if (opacity >= 1) {
                stafflineElement.style.removeProperty("opacity");
            } else {
                stafflineElement.style.opacity = opacityString;
            }
        }
    }

    private getStafflineElements(staffIndex: number, withinRoots?: ParentNode[]): SVGGElement[] {
        if (!Number.isFinite(staffIndex) || staffIndex < 0) {
            return [];
        }
        const selector: string = `.staffline[data-staff-index="${staffIndex}"]`;
        const elements: SVGGElement[] = [];
        const visitedRoots: Set<Element> = new Set();
        const seenElements: Set<SVGGElement> = new Set();
        const addMatches: (root: ParentNode) => void = (root: ParentNode): void => {
            const matches: NodeListOf<SVGGElement> = root.querySelectorAll<SVGGElement>(selector);
            for (const match of matches) {
                if (seenElements.has(match)) {
                    continue;
                }
                seenElements.add(match);
                elements.push(match);
            }
        };

        if (withinRoots) {
            for (const root of withinRoots) {
                addMatches(root);
            }
            return elements;
        }

        if (this.drawer?.Backends?.length) {
            for (const backend of this.drawer.Backends) {
                const rootElement: HTMLElement = backend.getRenderElement();
                if (!rootElement || visitedRoots.has(rootElement)) {
                    continue;
                }
                addMatches(rootElement);
                visitedRoots.add(rootElement);
            }
        }

        if (this.container && !visitedRoots.has(this.container)) {
            addMatches(this.container);
        }

        for (const systemGroup of this.drawer?.SystemGroups ?? []) {
            if (!systemGroup.isConnected) {
                addMatches(systemGroup);
            }
        }

        return elements;
    }

    private isValidStaffIndex(staffIndex: number): boolean {
        if (!Number.isFinite(staffIndex) || staffIndex < 0) {
            return false;
        }
        if (!this.sheet) {
            return true;
        }
        return staffIndex < this.sheet.getCompleteNumberOfStaves();
    }

    private clampOpacity(opacity: number, fallbackOpacity: number): number {
        const fallback: number = Number.isFinite(fallbackOpacity) ? fallbackOpacity : 1.0;
        const candidate: number = Number.isFinite(opacity) ? opacity : fallback;
        return Math.min(1, Math.max(0, candidate));
    }

    private reapplyStaffOpacityOverrides(withinRoots?: ParentNode[]): void {
        if (this.staffOpacityOverrides.size === 0) {
            return;
        }
        for (const [staffIndex, storedOpacity] of this.staffOpacityOverrides.entries()) {
            this.applyOpacityToStaffElements(staffIndex, storedOpacity, withinRoots);
        }
    }
}

/** What the steps of one load() pass along (see OpenSheetMusicDisplay.continueLoad()). Every field is required, so that
 *  no step can leave out isFileContent, whose missing value would allow a request. */
interface ILoadState {
    /** load()'s tempTitle: the title for a score without one. */
    tempTitle: string;
    /** The number of the load() the steps belong to (see OpenSheetMusicDisplay.loadGeneration). */
    generation: number;
    /** Whether a string content was read from a file (a Blob, an MXL archive, a URL's response) instead of given by the
     *  app: then it is never taken as a URL (see OpenSheetMusicDisplay.loadContent()). */
    isFileContent: boolean;
}

/** Options for {@link OpenSheetMusicDisplay.renderNext} (incremental, "system by system" rendering). */
export interface IRenderNextOptions {
    /** Target number of visual measures to advance per batch (a multi-rest counts as ONE -- it renders as a
     *  single GraphicalMeasure). Defaults to 8. This is a TARGET, not an exact count: a batch always ends on a
     *  whole music-system (line) boundary, so the system the measure count lands inside is deferred to the next
     *  batch (and a batch always renders at least one whole system). The measures actually drawn may therefore
     *  be somewhat fewer or more than this -- e.g. 8 lands part-way into a system, so only the complete systems
     *  before it are drawn; or if 8 doesn't fill one system, the layout extends until one whole system is ready.
     *  To see what was actually rendered, read the returned {@link IRenderNextResult}: `renderedMeasures` (total
     *  visual measures drawn so far, cumulative) and `lastRenderedMeasure` (the GraphicalMeasures at the last
     *  drawn measure position). Ignored when `systems` is set (and applicable). */
    measures?: number;
    /** How many whole music systems (lines) to render in this batch, instead of advancing by `measures`.
     *  Each batch then ends exactly on a system boundary. Takes precedence over `measures` when > 0.
     *  VERTICAL ONLY: a single horizontal staffline (RenderSingleHorizontalStaffline) is one system, so
     *  `systems` is ignored there and rendering falls back to `measures`. */
    systems?: number;
}

/** Progress returned by {@link OpenSheetMusicDisplay.renderNext}. Measure counts are visual (a multi-rest
 *  counts as one). */
export interface IRenderNextResult {
    /** True once the last measure to render has been rendered -- no more batches remain. The measures to render are the
     *  ones render() draws: the whole sheet, or the draw range (drawFromMeasureNumber, drawUpToMeasureNumber), up to the last
     *  system or page drawn (drawUpToSystemNumber, drawUpToPageNumber). */
    done: boolean;
    /** Visual measures rendered so far, cumulative across batches. */
    renderedMeasures: number;
    /** Total visual measures to render: those of the sheet, or of the draw range (see `done`). With drawUpToSystemNumber or
     *  drawUpToPageNumber, the measures after the last system or page drawn count until a batch's layout reaches it: where it
     *  ends depends on the layout. On a single horizontal staffline, the first batch lays out all of them. */
    totalMeasures: number;
    /** The last measure position rendered so far (highest measure index): all its GraphicalMeasures, one per
     *  staff/instrument (e.g. 3 for a voice + piano score). Empty if nothing has been rendered yet. They
     *  share one source measure -- reach it and its number/index via any element's `.parentSourceMeasure` and
     *  `.parentSourceMeasure.measureListIndex` (0-based). */
    lastRenderedMeasure: GraphicalMeasure[];
    /** The next measure position not yet rendered (the frontier): all its GraphicalMeasures (one per staff),
     *  or empty once `done`. Same `.parentSourceMeasure` / `.parentSourceMeasure.measureListIndex` accessors. */
    nextUnrenderedMeasure: GraphicalMeasure[];
}
