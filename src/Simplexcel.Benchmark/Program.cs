using BenchmarkDotNet.Running;
using System;

namespace Simplexcel.Benchmark
{
    class Program
    {
        static void Main(string[] args)
        {
            var summary = BenchmarkRunner.Run<SimplexcelBenchmark>();
        }
    }
}