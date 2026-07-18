// Parallelization hygiene (tested by convention, not by attribute):
// - This assembly MUST NOT call Environment.SetEnvironmentVariable or otherwise
//   mutate process-level state. DocxEditor.Tests/AssemblyInfo.cs is the only
//   assembly that mutates the environment (PATH), which is why it alone carries
//   [assembly: CollectionBehavior(DisableTestParallelization = true)].
// - Because this assembly never touches shared process state, it intentionally
//   keeps xUnit's default in-assembly parallelization enabled. Keep it that way:
//   if a future test here needs env mutation, move that test to DocxEditor.Tests.
