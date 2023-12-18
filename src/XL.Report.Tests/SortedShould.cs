// namespace XL.Report.Tests;
//
// public sealed class SortedShould
// {
//     [Test]
//     public void METHOD()
//     {
//         var sorted = new Sorted<int, Int>();
//
//         sorted.Add(10);
//         sorted.Add(11);
//         sorted.Add(12);
//         sorted.Add(13);
//         sorted.Add(14);
//         sorted.Add(9);
//     }
//
//     private readonly record struct Int(int Value) : IKeyed<int>
//     {
//         public static implicit operator Int(int value) => new(value);
//         public int Key => Value;
//     }
// }