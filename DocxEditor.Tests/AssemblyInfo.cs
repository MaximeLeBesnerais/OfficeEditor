using Xunit;

// Some tests intentionally mutate process-level environment variables (PATH) to
// exercise CLI fallback behavior. Disable parallel execution for this assembly
// so those mutations cannot race other tests in the same testhost process.
[assembly: CollectionBehavior(DisableTestParallelization = true)]
