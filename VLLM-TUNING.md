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

Before v0.6.0, the server truncated the API's 4096-dimensional response locally. It now requests 1024 dimensions for Qwen3-Embedding when the route supports it, with a fallback to the original behavior.

## Enabling reduced API responses

For a LiteLLM route using the OpenAI provider, allow the dimensions parameter for this model:

```yaml
model_list:
  - model_name: Qwen/Qwen3-Embedding-8B
    litellm_params:
      model: openai/Qwen/Qwen3-Embedding-8B
      api_base: http://172.17.0.1:8006/v1
      allowed_openai_params: ["dimensions"]
```

Preserve the other model settings and credentials. Restart the proxy to load the updated configuration.

If vLLM reports that Qwen3-Embedding does not support Matryoshka embeddings, add this startup argument, as in the Qwen project's vLLM example:

```text
--hf-overrides '{"is_matryoshka":true}'
```

This marks the model as MRL-capable; it does not change the weights. Keep the default API output at 4096 dimensions and request 1024 per call. A global 1024 output default would change native dimension detection and existing index identity.

In the tested Open WebUI -> LiteLLM -> vLLM route, both changes were required. Eight real-mail chunks totaling 10,451 characters produced about 703 KB at 4096 dimensions and 174 KB at 1024 dimensions (about 75% less response data). After warmup, repeated identical-input requests had medians of 210 ms and 113 ms respectively. This is a small repeated-input experiment, not an estimate of full indexing speed.

The minimum cosine similarity between the API's reduced vectors and renormalized prefixes of its full vectors was above 0.999999999999999. A 28-query known-target evaluation across 12 real-mail topics, including date-filter cases, produced identical aggregate ranking metrics with both response modes: 18/28 at rank 1, 22/28 in the top 5, and MRR@10 of 0.7176. The labels are not exhaustive relevance judgments.

Shorter responses reduce transfer and JSON processing; they do not remove the model's main forward-pass computation. Existing 1024-dimensional indexes can be retained when the reduced vectors match the previous normalized prefixes.

## References

- [vLLM optimization](https://docs.vllm.ai/en/stable/configuration/optimization/)
- [Embedding and dimension settings](https://docs.vllm.ai/en/stable/models/pooling_models/embed/)
- [Qwen3-Embedding](https://github.com/QwenLM/Qwen3-Embedding)
