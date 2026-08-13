using md2visio.struc.graph;

namespace md2visio.Tests.Graph;

public sealed class CompoundGraphLayoutTests
{
    [Fact]
    public void Subgraph_InheritsParentDirection()
    {
        var parent = new md2visio.struc.graph.Graph { Direction = "TB" };

        var child = new GSubgraph(parent);

        Assert.Equal("TB", child.Direction);
    }

    [Fact]
    public void Calculate_KeepsGroupsDisjointAndContainsTheirNodes()
    {
        LayoutNode[] nodes =
        [
            new("prospect", null, 1.2, 0.5, 0),
            new("buy", null, 1.4, 0.5, 1),
            new("m1", "marketing", 1.6, 0.5, 2),
            new("m2", "marketing", 1.3, 0.5, 3),
            new("s1", "sales", 1.4, 0.5, 4),
            new("s2", "sales", 1.5, 0.5, 5)
        ];
        LayoutGroup[] groups =
        [
            new("marketing", null, "TB", 0),
            new("sales", null, "TB", 1)
        ];
        LayoutEdge[] edges =
        [
            new("prospect", "buy"),
            new("prospect", "m1"),
            new("m1", "m2"),
            new("m2", "s1"),
            new("s1", "s2")
        ];

        var result = CompoundGraphLayout.Calculate(nodes, groups, edges, "TB");

        AssertContains(result.Groups["marketing"], result.Nodes["m1"]);
        AssertContains(result.Groups["marketing"], result.Nodes["m2"]);
        AssertContains(result.Groups["sales"], result.Nodes["s1"]);
        AssertContains(result.Groups["sales"], result.Nodes["s2"]);
        Assert.False(Overlaps(result.Groups["marketing"], result.Groups["sales"]));
        Assert.True(result.Nodes["prospect"].CenterY > result.Nodes["buy"].CenterY);
    }

    [Fact]
    public void Calculate_CollapsesCyclesIntoOneRankWithoutCollisions()
    {
        LayoutNode[] nodes =
        [
            new("a", null, 1, 0.5, 0),
            new("b", null, 1, 0.5, 1),
            new("c", null, 1, 0.5, 2)
        ];
        LayoutEdge[] edges = [new("a", "b"), new("b", "a"), new("b", "c")];

        var result = CompoundGraphLayout.Calculate(nodes, [], edges, "TB");

        Assert.Equal(result.Nodes["a"].CenterY, result.Nodes["b"].CenterY, precision: 8);
        Assert.False(Overlaps(result.Nodes["a"], result.Nodes["b"]));
        Assert.True(result.Nodes["c"].CenterY < result.Nodes["b"].CenterY);
    }

    private static void AssertContains(LayoutRect outer, LayoutRect inner)
    {
        Assert.True(inner.CenterX - inner.Width / 2 >= outer.CenterX - outer.Width / 2);
        Assert.True(inner.CenterX + inner.Width / 2 <= outer.CenterX + outer.Width / 2);
        Assert.True(inner.CenterY - inner.Height / 2 >= outer.CenterY - outer.Height / 2);
        Assert.True(inner.CenterY + inner.Height / 2 <= outer.CenterY + outer.Height / 2);
    }

    private static bool Overlaps(LayoutRect a, LayoutRect b) =>
        Math.Abs(a.CenterX - b.CenterX) < (a.Width + b.Width) / 2 &&
        Math.Abs(a.CenterY - b.CenterY) < (a.Height + b.Height) / 2;
}
