using System.Runtime.CompilerServices;

// The Excel test project asserts on a few internal seams (e.g. the diff cell-cap
// constant) so its expectations track the implementation instead of hard-coding
// magic numbers that drift.
[assembly: InternalsVisibleTo("AIOffice.Excel.Tests")]
