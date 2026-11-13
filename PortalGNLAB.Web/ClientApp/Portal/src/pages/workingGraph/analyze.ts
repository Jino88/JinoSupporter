import { toolSend } from "../../toolsApi";
import type { Payload } from "../graphMaker/Charts";
import type { Table } from "../graphMaker/data";
import {
  buildRequests,
  graphDefaults,
  loadedOptions,
  type GraphOptions,
} from "../graphMaker/options";

export interface AnalyzeResult {
  kind: number;
  title: string;
  reason: string;
  options: Partial<GraphOptions>;
  model: string;
  elapsedSeconds: number;
}

export interface DataInput {
  text: string;
  delimiter: string;
  headerRowNumber: number;
  headerNotUse: boolean;
}

/** The AI's pick laid over the defaults and the columns of the pasted table. */
export function mergeAnalysis(table: Table, result: AnalyzeResult): GraphOptions {
  return {
    ...graphDefaults,
    ...loadedOptions(table, graphDefaults),
    ...result.options,
  };
}

export async function renderGraphs(
  input: DataInput,
  kind: number,
  options: GraphOptions,
  signal?: AbortSignal,
): Promise<Payload[]> {
  const payloads: Payload[] = [];
  for (const request of buildRequests(kind, options)) {
    payloads.push(
      await toolSend<Payload>(
        "working-graph/render",
        { ...input, request },
        "POST",
        signal,
      ),
    );
  }
  return payloads;
}

/** Asks the server's CLI which graph fits the data, then draws it right away. */
export async function analyzeAndRender(
  input: DataInput,
  table: Table,
  signal?: AbortSignal,
  onAnalyzed?: (result: AnalyzeResult, options: GraphOptions) => void,
) {
  const result = await toolSend<AnalyzeResult>(
    "working-graph/analyze",
    input,
    "POST",
    signal,
  );
  const options = mergeAnalysis(table, result);
  onAnalyzed?.(result, options);
  const payloads = await renderGraphs(input, result.kind, options, signal);
  return { result, options, payloads };
}
