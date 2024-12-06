namespace XL.Report;

internal sealed partial class StreamSheetWindow
{
    // ReSharper disable ParameterOnlyUsedForPreconditionCheck.Local
    private sealed class ReductionStage(StreamSheetWindow window, int index) : IDisposable
    // ReSharper restore ParameterOnlyUsedForPreconditionCheck.Local
    {
        public void Dispose()
        {
            if (window.reductions.Count != index + 1)
            {
                throw new InvalidOperationException();
            }

            window.reductions.Pop();
        }
    }
}