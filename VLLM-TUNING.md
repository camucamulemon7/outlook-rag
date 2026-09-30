# vLLM measurements and tuning

These are measurements from one local environment, not universal performance recommendations.

Environment: vLLM 0.30.0, Qwen3-Embedding-8B, FP8, RTX 5060 Ti 16GB. The application route was Open WebUI -> LiteLLM -> vLLM. The MCP server submitted up to 8 chunks per request with 2 requests in flight; this is separate from Open WebUI's built-in RAG batching settings.

## Where the time went

For 8 chunks totaling 12,433 characters, the configured route took 1.40-1.43 seconds, compared with 1.35-1.37 seconds when calling vLLM directly. Each route was tested twice in direct/configured/configured/direct order. Unique prefixes reduced identical-input cache reuse. The sample is too small to establish a statistically reliable difference, and the application route was retained.

With `max-num-seqs=4`, an 8-input embedding request produced 8 sequences, of which at most 4 could run concurrently. GPU utilization reached 100%, with no recorded preemptions. Per-sequence queue and inference metrics are cumulative and can exceed HTTP wall time when added together. GPU work and sequence waiting appeared to dominate proxy overhead in these measurements.

## Comparing max-num-seqs=4 and 64

After the user increased `max-num-seqs` to 64, the running process was checked. `max-num-batched-tokens=8192`, FP8, `enforce-eager`, and `gpu-memory-utilization=0.85` were retained.

The MCP settings were 8 chunks per request, 2 concurrent requests, 1024 stored dimensions, and enhanced text cleanup. The same 40 emails required 87 new embeddings in an empty database.

| Condition | Total time | Emails/second |
| --- | --- | --- |
| max-num-seqs=4, earlier single measurement | 11.013 s | 3.63 |
| max-num-seqs=64, same document inputs | 11.263 s | 3.55 |
| max-num-seqs=64, unique input prefixes, run 1 | 11.841 s | 3.38 |
| max-num-seqs=64, unique input prefixes, run 2 | 11.400 s | 3.51 |

The extra prefixes reduced reuse of identical inputs but did not disable all prefix caching; they also added tokens. The first two rows are the closest input comparison, but the baseline was measured earlier, not in an alternating experiment.

No speed improvement was observed. GPU utilization reached 100%, no additional preemptions were recorded, and all 87 inputs completed. GPU compute and the batched-token limit may still constrain throughput. `max-num-seqs` sets an upper bound; it does not increase GPU processing capacity.

## Further experiments

Change one setting at a time and compare total completion time, tokens/second, queue time, memory usage, and errors on the same input:

- Compare `max-num-batched-tokens=8192` with `16384` if memory permits.
- Compare eager execution with compilation/CUDA Graphs where supported by the pooling implementation and vLLM version.
- Recheck request batching and concurrency; increasing MCP concurrency from 2 to 4 did not improve earlier measurements.

These experiments were not applied as part of the measurements above. Reducing `max-model-len` does not itself shorten the actual input.

## Reducing input work

On 251 emails, enhanced cleanup reduced embedding input from 1,656,336 to 775,417 characters (about 53%) and chunks from 1,207 to 591 (about 51%). In a separate 40-email comparison, indexing time dropped from 22.60 to 11.01 seconds. These are small local samples and may be affected by caching and load.

The current server truncates the API's 4096-dimensional response to 1024 dimensions and renormalizes it. This reduces local vector storage and search work, but does not reduce GPU inference or API response size. Requesting 1024 dimensions directly from vLLM would require separate API compatibility checks.

## References

- [vLLM optimization](https://docs.vllm.ai/en/stable/configuration/optimization/)
- [Embedding and dimension settings](https://docs.vllm.ai/en/stable/models/pooling_models/embed/)
- [Qwen3-Embedding](https://github.com/QwenLM/Qwen3-Embedding)
