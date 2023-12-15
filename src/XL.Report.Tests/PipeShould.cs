using System.Buffers;
using System.IO.Pipelines;

namespace XL.Report.Tests;

public sealed class PipeShould
{
    [Test]
    public async Task METHOD()
    {
        var pipe = new Pipe();
        var buffer = new byte[1024];
        for (var i = 0; i < 1024; i++)
        {
            // for (var j = 0; j < 1024; j++)
            {
                pipe.Writer.Write(buffer);
            }
        }

        var result = pipe.Writer.FlushAsync();
        
        pipe.Reader.Complete(new Exception("B"));
        pipe.Reader.Complete(new Exception("A"));
        // pipe.Reader.Complete();

        var flushResult = await result;
    }
}