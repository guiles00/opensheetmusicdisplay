import { expect } from "chai";
import { OpenSheetMusicDisplay } from "../../src/OpenSheetMusicDisplay/OpenSheetMusicDisplay";
import { generateLargePianoScore, ILargeScoreFixtureOptions } from "./LargeScoreFixtures";

interface IDistribution {
    count: number;
    total: number;
    p50: number;
    p95: number;
    max: number;
}

interface IBenchmarkWindow {
    __karma__?: { config?: { args?: string[] } };
    gc?: () => void;
}

const benchmarkWindow: IBenchmarkWindow = window as unknown as IBenchmarkWindow;
const benchmarkArgs: string[] = benchmarkWindow.__karma__?.config?.args ?? [];
const benchmarkEnabled: boolean = benchmarkArgs.includes("bench");
const describeBenchmark: Mocha.SuiteFunction | Mocha.PendingSuiteFunction = benchmarkEnabled ? describe : describe.skip;
const benchmarkLabel: string = benchmarkArgs.find(arg => arg.startsWith("label="))?.slice("label=".length) ?? "current";

function distribution(samples: number[]): IDistribution {
    const sorted: number[] = [...samples].sort((a, b): number => a - b);
    const pick: (quantile: number) => number = (quantile: number): number =>
        sorted.length ? sorted[Math.min(sorted.length - 1, Math.floor(quantile * sorted.length))] : 0;
    const round: (value: number) => number = (value: number): number => Math.round(value * 1000) / 1000;
    return {
        count: sorted.length,
        total: round(sorted.reduce((sum, value): number => sum + value, 0)),
        p50: round(pick(0.5)),
        p95: round(pick(0.95)),
        max: round(sorted[sorted.length - 1] ?? 0)
    };
}

function heapMegabytes(): number | undefined {
    benchmarkWindow.gc?.();
    const memory: { usedJSHeapSize: number } | undefined = (performance as unknown as { memory?: { usedJSHeapSize: number } }).memory;
    return memory ? Math.round(memory.usedJSHeapSize / 1024 / 102.4) / 10 : undefined;
}

function mount(): { scrollElement: HTMLDivElement, container: HTMLDivElement } {
    const scrollElement: HTMLDivElement = document.createElement("div");
    scrollElement.style.width = "1000px";
    scrollElement.style.height = "800px";
    scrollElement.style.overflow = "auto";
    const container: HTMLDivElement = document.createElement("div");
    container.style.width = "1000px";
    scrollElement.appendChild(container);
    document.body.appendChild(scrollElement);
    return { scrollElement, container };
}

function time(action: () => void): number {
    const start: number = performance.now();
    action();
    return performance.now() - start;
}

function retainedSvgNodeCount(osmd: OpenSheetMusicDisplay): number {
    let count: number = 0;
    for (const group of osmd.Drawer.SystemGroups) {
        count += group.getElementsByTagName("*").length;
    }
    return count;
}

async function benchmarkFixture(name: string, fixture: ILargeScoreFixtureOptions): Promise<Record<string, unknown>> {
    const xml: string = generateLargePianoScore(fixture);
    const heapBefore: number | undefined = heapMegabytes();

    const full: { scrollElement: HTMLDivElement, container: HTMLDivElement } = mount();
    const fullOsmd: OpenSheetMusicDisplay = new OpenSheetMusicDisplay(full.container, { autoResize: false });
    await fullOsmd.load(xml);
    const fullRenderMs: number = time(() => fullOsmd.render());
    const fullDomNodes: number = full.container.getElementsByTagName("*").length;
    full.scrollElement.remove();

    const { scrollElement, container } = mount();
    const osmd: OpenSheetMusicDisplay = new OpenSheetMusicDisplay(container, { autoResize: false });
    const loadStart: number = performance.now();
    await osmd.load(xml);
    const loadMs: number = performance.now() - loadStart;
    osmd.enableSystemVirtualization({ scrollElement, overscanViewports: 1 });
    const virtualizedRenderMs: number = time(() => osmd.renderVirtualized({ initialSystems: 3 }));
    const totalSystems: number = osmd.SystemVirtualizationStats.totalSystems;

    const firstPassUpdates: number[] = [];
    const materializationPerSystem: number[] = [];
    const step: number = scrollElement.clientHeight / 2;
    for (let top: number = 0; top <= scrollElement.scrollHeight; top += step) {
        scrollElement.scrollTop = top;
        const before: number = osmd.SystemVirtualizationStats.materializedSystems;
        const duration: number = time(() => osmd.updateSystemVirtualization());
        firstPassUpdates.push(duration);
        const drawn: number = osmd.SystemVirtualizationStats.materializedSystems - before;
        if (drawn > 0) {
            materializationPerSystem.push(duration / drawn);
        }
    }
    const materializedAfterScroll: number = osmd.SystemVirtualizationStats.materializedSystems;

    const viewportOnlyUpdates: number[] = [];
    for (let top: number = scrollElement.scrollHeight; top >= 0; top -= step) {
        scrollElement.scrollTop = top;
        viewportOnlyUpdates.push(time(() => osmd.updateSystemVirtualization()));
    }

    const readAheadHide: number[] = [];
    const measureCount: number = osmd.Sheet.SourceMeasures.length;
    for (let measureIndex: number = 0; measureIndex < measureCount; measureIndex++) {
        for (let beat: number = 0; beat < 3; beat++) {
            readAheadHide.push(time(() => osmd.hideReadAheadMeasureBeatRange(measureIndex, beat, 1, 0)));
        }
    }
    const readAheadResetMs: number = time(() => osmd.resetReadAheadOpacity());

    const result: Record<string, unknown> = {
        label: benchmarkLabel,
        fixture: name,
        measures: measureCount,
        totalSystems,
        loadMs: Math.round(loadMs),
        fullRenderMs: Math.round(fullRenderMs),
        virtualizedRenderMs: Math.round(virtualizedRenderMs),
        firstPassUpdateMs: distribution(firstPassUpdates),
        materializationPerSystemMs: distribution(materializationPerSystem),
        viewportOnlyUpdateMs: distribution(viewportOnlyUpdates),
        readAheadHideMs: distribution(readAheadHide),
        readAheadResetMs: Math.round(readAheadResetMs * 1000) / 1000,
        materializedAfterScroll,
        attachedAfterScroll: osmd.SystemVirtualizationStats.attachedSystems,
        liveDomNodes: container.getElementsByTagName("*").length,
        retainedSvgNodes: retainedSvgNodeCount(osmd),
        fullRenderDomNodes: fullDomNodes,
        heapBeforeMb: heapBefore,
        heapAfterScrollMb: heapMegabytes()
    };
    osmd.disableSystemVirtualization();
    scrollElement.remove();
    return result;
}

const FIXTURES: [string, ILargeScoreFixtureOptions][] = [
    ["piano-20", { measures: 20, graceEvery: 8, ornamentEvery: 6, tieEvery: 4, fingeringEvery: 4 }],
    ["piano-200", { measures: 200, graceEvery: 8, ornamentEvery: 6, tieEvery: 4, fingeringEvery: 4, meterChangeEvery: 32 }],
    ["piano-500", { measures: 500, graceEvery: 8, ornamentEvery: 6, tieEvery: 4, fingeringEvery: 4, meterChangeEvery: 32 }],
    ["piano-500-breaks", { measures: 500, graceEvery: 8, tieEvery: 4, systemBreakEvery: 4 }]
];

describeBenchmark("Large score benchmark", function (): void {
    this.timeout(30 * 60 * 1000);

    for (const [name, fixture] of FIXTURES) {
        it(name, async () => {
            const result: Record<string, unknown> = await benchmarkFixture(name, fixture);
            console.log(`BENCH ${JSON.stringify(result)}`);
            expect(result.materializedAfterScroll).to.equal(result.totalSystems);
        });
    }
});
