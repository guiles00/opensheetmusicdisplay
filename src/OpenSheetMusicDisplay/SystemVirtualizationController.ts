export interface ISystemVirtualizationOptions {
    /** The element that clips and scrolls the score. Defaults to window. */
    scrollElement?: HTMLElement | Window;
    /** Extra viewport heights kept mounted above and below the visible area. Defaults to 1. */
    overscanViewports?: number;
}

export interface ISystemVirtualizationStats {
    totalSystems: number;
    materializedSystems: number;
    attachedSystems: number;
    detachedSystems: number;
    unmaterializedSystems: number;
}

export type SystemLifecycleEventType = "materialized" | "attached" | "detached";

/** Systems whose SVG was drawn for the first time, reinserted, or removed from the live DOM in one update. */
export interface ISystemLifecycleEvent {
    type: SystemLifecycleEventType;
    keys: string[];
    roots: SVGGElement[];
}

export type SystemLifecycleListener = (event: ISystemLifecycleEvent) => void;

export interface IVirtualSystemDescriptor {
    key: string;
    svg: SVGSVGElement;
    top: number;
    bottom: number;
}

interface VirtualizedSystem {
    group: SVGGElement;
    anchor: Comment;
    svg: SVGSVGElement;
    top: number;
    bottom: number;
    attached: boolean;
}

interface IndexedSystem {
    key: string;
    top: number;
    bottom: number;
}

interface SvgSystemIndex {
    systems: IndexedSystem[];
    maxBottomThrough: number[];
}

type MaterializeSystems = (keys: string[]) => SVGGElement[] | void;

interface VirtualizationViewport {
    top: number;
    bottom: number;
    height: number;
}

/**
 * Keeps only nearby music-system SVG groups in the live DOM. The groups themselves are retained and
 * reinserted, rather than recreated, so colors, opacity, listeners, and application-owned state survive.
 */
export class SystemVirtualizationController {
    private readonly container: HTMLElement;
    private readonly systems: Map<string, VirtualizedSystem> = new Map<string, VirtualizedSystem>();
    private readonly expectedSystems: Map<string, IVirtualSystemDescriptor> = new Map<string, IVirtualSystemDescriptor>();
    private enabled: boolean = false;
    private target: HTMLElement | Window | undefined;
    private overscanViewports: number = 1;
    private frameRequest: number | undefined;
    private materializeFrameRequest: number | undefined;
    private readonly pendingMaterializationKeys: Set<string> = new Set<string>();
    private materializeSystems: MaterializeSystems | undefined;
    private systemIndex: Map<SVGSVGElement, SvgSystemIndex> | undefined;
    private readonly attachedKeys: Set<string> = new Set<string>();
    private readonly lifecycleListeners: Set<SystemLifecycleListener> = new Set<SystemLifecycleListener>();

    public constructor(container: HTMLElement) {
        this.container = container;
    }

    public enable(options?: ISystemVirtualizationOptions): void {
        if (typeof window === "undefined") {
            return;
        }
        this.disableListeners();
        this.enabled = true;
        this.target = options?.scrollElement ?? window;
        this.overscanViewports = Math.max(0, options?.overscanViewports ?? 1);
        this.target.addEventListener("scroll", this.scheduleUpdate, { passive: true });
        window.addEventListener("resize", this.scheduleUpdate, { passive: true });
        this.refresh();
    }

    public disable(restore: boolean = true): void {
        this.enabled = false;
        this.disableListeners();
        if (restore) {
            const attached: [string, SVGGElement][] = [];
            for (const [key, system] of this.systems) {
                if (this.attach(key, system)) {
                    attached.push([key, system.group]);
                }
            }
            this.emit("attached", attached);
        }
    }

    /** Register the full laid-out score, including systems that have not been drawn yet. */
    public configureExpectedSystems(
        systems: IVirtualSystemDescriptor[],
        materialize: MaterializeSystems
    ): void {
        this.expectedSystems.clear();
        for (const system of systems) {
            this.expectedSystems.set(system.key, system);
        }
        this.systemIndex = undefined;
        this.materializeSystems = materialize;
        this.refresh();
    }

    /** Subscribe to system materialization, attach and detach; returns an unsubscribe function. */
    public addLifecycleListener(listener: SystemLifecycleListener): () => void {
        this.lifecycleListeners.add(listener);
        return (): void => {
            this.lifecycleListeners.delete(listener);
        };
    }

    private emit(type: SystemLifecycleEventType, systems: [string, SVGGElement][]): void {
        if (systems.length === 0 || this.lifecycleListeners.size === 0) {
            return;
        }
        const event: ISystemLifecycleEvent = {
            type,
            keys: systems.map(([key]) => key),
            roots: systems.map(([, root]) => root)
        };
        for (const listener of Array.from(this.lifecycleListeners)) {
            listener(event);
        }
    }

    /** Drop references before OSMD destroys/replaces a backend. */
    public invalidate(): void {
        if (this.frameRequest !== undefined && typeof window !== "undefined") {
            window.cancelAnimationFrame(this.frameRequest);
        }
        this.frameRequest = undefined;
        if (this.materializeFrameRequest !== undefined && typeof window !== "undefined") {
            window.cancelAnimationFrame(this.materializeFrameRequest);
        }
        this.materializeFrameRequest = undefined;
        this.pendingMaterializationKeys.clear();
        this.systems.clear();
        this.expectedSystems.clear();
        this.attachedKeys.clear();
        this.systemIndex = undefined;
        this.materializeSystems = undefined;
    }

    /** Discover newly rendered systems, then apply the current viewport window. */
    public refresh(): void {
        if (!this.enabled || typeof document === "undefined") {
            return;
        }
        this.emit("materialized", this.discoverRenderedSystems());
        this.updateNow();
        // Layout may not have settled yet right after this DOM mutation (see the null-CTM handling in
        // updateNow()); schedule one more evaluation next frame to catch what couldn't be measured above.
        this.scheduleUpdate();
    }

    private discoverRenderedSystems(): [string, SVGGElement][] {
        return this.registerSystemGroups(Array.from(
            this.container.querySelectorAll<SVGGElement>("g.osmd-system[data-osmd-system-key]")
        ));
    }

    private registerSystemGroups(groups: SVGGElement[]): [string, SVGGElement][] {
        const registered: [string, SVGGElement][] = [];
        for (const group of groups) {
            const key: string = group.dataset.osmdSystemKey;
            if (!key || this.systems.has(key)) {
                continue;
            }
            const svg: SVGSVGElement = group.ownerSVGElement;
            if (!svg || !group.parentNode) {
                continue;
            }
            const top: number = Number.parseFloat(group.dataset.osmdSystemTop ?? "");
            const bottom: number = Number.parseFloat(group.dataset.osmdSystemBottom ?? "");
            if (!Number.isFinite(top) || !Number.isFinite(bottom)) {
                continue;
            }
            const anchor: Comment = document.createComment(`osmd-system:${key}`);
            group.parentNode.insertBefore(anchor, group);
            this.systems.set(key, { group, anchor, svg, top, bottom, attached: true });
            this.attachedKeys.add(key);
            if (!this.expectedSystems.has(key)) {
                this.systemIndex = undefined;
            }
            registered.push([key, group]);
        }
        return registered;
    }

    private materialize(keys: string[]): void {
        const groups: SVGGElement[] | void = this.materializeSystems?.(keys);
        this.emit("materialized", groups ? this.registerSystemGroups(groups) : this.discoverRenderedSystems());
    }

    /** Systems per SVG in vertical order, so a viewport window is found by binary search. */
    private getSystemIndex(): Map<SVGSVGElement, SvgSystemIndex> {
        if (this.systemIndex) {
            return this.systemIndex;
        }
        const bySvg: Map<SVGSVGElement, IndexedSystem[]> = new Map<SVGSVGElement, IndexedSystem[]>();
        const add: (svg: SVGSVGElement, system: IndexedSystem) => void = (svg: SVGSVGElement, system: IndexedSystem): void => {
            const systems: IndexedSystem[] = bySvg.get(svg);
            if (systems) {
                systems.push(system);
            } else {
                bySvg.set(svg, [system]);
            }
        };
        for (const expected of this.expectedSystems.values()) {
            add(expected.svg, { key: expected.key, top: expected.top, bottom: expected.bottom });
        }
        for (const [key, system] of this.systems) {
            if (!this.expectedSystems.has(key)) {
                add(system.svg, { key, top: system.top, bottom: system.bottom });
            }
        }
        this.systemIndex = new Map<SVGSVGElement, SvgSystemIndex>();
        for (const [svg, systems] of bySvg) {
            systems.sort((a, b): number => a.top - b.top || a.bottom - b.bottom);
            const maxBottomThrough: number[] = [];
            let maxBottom: number = Number.NEGATIVE_INFINITY;
            for (const system of systems) {
                maxBottom = Math.max(maxBottom, system.bottom);
                maxBottomThrough.push(maxBottom);
            }
            this.systemIndex.set(svg, { systems, maxBottomThrough });
        }
        return this.systemIndex;
    }

    private systemsIntersecting(
        index: SvgSystemIndex,
        matrix: DOMMatrix,
        minY: number,
        maxY: number
    ): { system: IndexedSystem, top: number, bottom: number }[] {
        const sorted: boolean = matrix.d > 0;
        let first: number = 0;
        if (sorted) {
            const localMinY: number = (minY - matrix.f) / matrix.d - 1;
            let high: number = index.systems.length;
            while (first < high) {
                const mid: number = Math.floor((first + high) / 2);
                if (index.maxBottomThrough[mid] < localMinY) {
                    first = mid + 1;
                } else {
                    high = mid;
                }
            }
        }
        const hits: { system: IndexedSystem, top: number, bottom: number }[] = [];
        for (let i: number = first; i < index.systems.length; i++) {
            const system: IndexedSystem = index.systems[i];
            const top: number = new DOMPoint(0, system.top).matrixTransform(matrix).y;
            if (sorted && top > maxY) {
                break;
            }
            const bottom: number = new DOMPoint(0, system.bottom).matrixTransform(matrix).y;
            if (bottom >= minY && top <= maxY) {
                hits.push({ system, top, bottom });
            }
        }
        return hits;
    }

    public updateNow(): void {
        if (!this.enabled || (this.systems.size === 0 && this.expectedSystems.size === 0)) {
            return;
        }
        const viewport: VirtualizationViewport = this.getViewport();
        if (viewport.height <= 0) {
            return;
        }
        const minY: number = viewport.top - viewport.height * this.overscanViewports;
        const maxY: number = viewport.bottom + viewport.height * this.overscanViewports;
        const missingVisibleKeys: string[] = [];
        const missingOverscanKeys: string[] = [];
        const inRange: Set<string> = new Set<string>();
        const measuredSvgs: Set<SVGSVGElement> = new Set<SVGSVGElement>();
        for (const [svg, index] of this.getSystemIndex()) {
            // getScreenCTM() forces a synchronous layout flush, so read it once per SVG before any DOM writes.
            const matrix: DOMMatrix = svg.getScreenCTM();
            if (!matrix) {
                // Not measurable right now (e.g. mid-backend-swap): leave these systems as they are; the
                // follow-up frame scheduled in refresh() re-evaluates them once geometry is available.
                continue;
            }
            measuredSvgs.add(svg);
            for (const hit of this.systemsIntersecting(index, matrix, minY, maxY)) {
                inRange.add(hit.system.key);
                if (this.systems.has(hit.system.key)) {
                    continue;
                }
                if (hit.bottom >= viewport.top && hit.top <= viewport.bottom) {
                    missingVisibleKeys.push(hit.system.key);
                } else {
                    missingOverscanKeys.push(hit.system.key);
                }
            }
        }
        // Systems intersecting the real viewport are urgent: draw them now so cursor jumps and fast
        // scrolling never expose a blank line. Overscan is speculative and is spread over later frames.
        if (missingVisibleKeys.length > 0 && this.materializeSystems) {
            this.materialize(missingVisibleKeys);
        }
        this.pendingMaterializationKeys.clear();
        for (const key of missingOverscanKeys) {
            if (!this.systems.has(key)) {
                this.pendingMaterializationKeys.add(key);
            }
        }
        this.scheduleNextMaterialization();

        const attached: [string, SVGGElement][] = [];
        for (const key of inRange) {
            const system: VirtualizedSystem = this.systems.get(key);
            if (system && this.attach(key, system)) {
                attached.push([key, system.group]);
            }
        }
        const detached: [string, SVGGElement][] = [];
        for (const key of Array.from(this.attachedKeys)) {
            const system: VirtualizedSystem = this.systems.get(key);
            if (system && !inRange.has(key) && measuredSvgs.has(system.svg) && this.detach(key, system)) {
                detached.push([key, system.group]);
            }
        }
        this.emit("attached", attached);
        this.emit("detached", detached);
    }

    public get stats(): ISystemVirtualizationStats {
        const attachedSystems: number = this.attachedKeys.size;
        return {
            totalSystems: Math.max(this.expectedSystems.size, this.systems.size),
            materializedSystems: this.systems.size,
            attachedSystems,
            detachedSystems: this.systems.size - attachedSystems,
            unmaterializedSystems: Math.max(0, this.expectedSystems.size - this.systems.size)
        };
    }

    private readonly scheduleUpdate: () => void = (): void => {
        if (!this.enabled || this.frameRequest !== undefined) {
            return;
        }
        this.frameRequest = window.requestAnimationFrame((): void => {
            this.frameRequest = undefined;
            this.updateNow();
        });
    };

    private scheduleNextMaterialization(): void {
        if (!this.enabled || !this.materializeSystems || this.pendingMaterializationKeys.size === 0 ||
            this.materializeFrameRequest !== undefined) {
            return;
        }
        this.materializeFrameRequest = window.requestAnimationFrame((): void => {
            this.materializeFrameRequest = undefined;
            if (!this.enabled || !this.materializeSystems) {
                return;
            }
            const iterator: Iterator<string> = this.pendingMaterializationKeys.values();
            const next: IteratorResult<string> = iterator.next();
            if (next.done) {
                return;
            }
            this.pendingMaterializationKeys.delete(next.value);
            if (!this.systems.has(next.value)) {
                this.materialize([next.value]);
            }
            this.scheduleNextMaterialization();
        });
    }

    private getViewport(): VirtualizationViewport {
        if (!this.target || this.target === window) {
            return { top: 0, bottom: window.innerHeight, height: window.innerHeight };
        }
        const rect: DOMRect = (this.target as HTMLElement).getBoundingClientRect();
        return { top: rect.top, bottom: rect.bottom, height: rect.height };
    }

    private attach(key: string, system: VirtualizedSystem): boolean {
        if (system.attached) {
            return false;
        }
        system.anchor.parentNode?.insertBefore(system.group, system.anchor.nextSibling);
        system.attached = true;
        this.attachedKeys.add(key);
        return true;
    }

    private detach(key: string, system: VirtualizedSystem): boolean {
        if (!system.attached) {
            return false;
        }
        system.group.remove();
        system.attached = false;
        this.attachedKeys.delete(key);
        return true;
    }

    private disableListeners(): void {
        if (typeof window === "undefined") {
            return;
        }
        this.target?.removeEventListener("scroll", this.scheduleUpdate);
        window.removeEventListener("resize", this.scheduleUpdate);
        if (this.frameRequest !== undefined) {
            window.cancelAnimationFrame(this.frameRequest);
            this.frameRequest = undefined;
        }
        if (this.materializeFrameRequest !== undefined) {
            window.cancelAnimationFrame(this.materializeFrameRequest);
            this.materializeFrameRequest = undefined;
        }
        this.pendingMaterializationKeys.clear();
    }
}
