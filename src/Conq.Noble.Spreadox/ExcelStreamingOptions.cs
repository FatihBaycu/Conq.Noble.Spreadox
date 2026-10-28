namespace Conq.Noble.Spreadox
{
    /// <summary>
    /// Options for <see cref="ExcelStreamingWorkbook"/>.
    /// </summary>
    public sealed class ExcelStreamingOptions
    {
        /// <summary>
        /// When <c>true</c> (default), string cells use <c>inlineStr</c> so the shared-string table
        /// stays empty — best for unique / high-cardinality text and lower peak memory.
        /// When <c>false</c>, strings go through SST (smaller files when many repeats; SST grows in memory).
        /// </summary>
        public bool UseInlineStrings { get; set; } = true;
    }
}
