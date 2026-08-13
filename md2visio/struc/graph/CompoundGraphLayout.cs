namespace md2visio.struc.graph;

internal sealed record LayoutNode(string Id, string? ParentId, double Width, double Height, int Order);
internal sealed record LayoutGroup(string Id, string? ParentId, string Direction, int Order);
internal sealed record LayoutEdge(string FromId, string ToId);
internal sealed record LayoutRect(double CenterX, double CenterY, double Width, double Height);

internal sealed class CompoundGraphLayoutResult
{
    public Dictionary<string, LayoutRect> Nodes { get; } = new(StringComparer.Ordinal);
    public Dictionary<string, LayoutRect> Groups { get; } = new(StringComparer.Ordinal);
}

internal static class CompoundGraphLayout
{
    private const string RootId = "$root";
    private const double ItemGap = 0.65;
    private const double RankGap = 0.9;
    private const double GroupPadding = 0.45;
    private const double GroupTitleHeight = 0.3;

    private sealed class Box
    {
        public required string Id { get; init; }
        public required bool IsGroup { get; init; }
        public required int Order { get; init; }
        public double Width { get; set; }
        public double Height { get; set; }
        public double X { get; set; }
        public double Y { get; set; }
        public List<Box> Children { get; } = [];
    }

    public static CompoundGraphLayoutResult Calculate(
        IReadOnlyCollection<LayoutNode> nodes,
        IReadOnlyCollection<LayoutGroup> groups,
        IReadOnlyCollection<LayoutEdge> edges,
        string rootDirection)
    {
        var nodesById = nodes.ToDictionary(n => n.Id, StringComparer.Ordinal);
        var groupsById = groups.ToDictionary(g => g.Id, StringComparer.Ordinal);
        var childrenByParent = groups
            .GroupBy(g => g.ParentId ?? RootId)
            .ToDictionary(g => g.Key, g => g.OrderBy(x => x.Order).ToList(), StringComparer.Ordinal);

        string ParentOfNode(string nodeId) => nodesById[nodeId].ParentId ?? RootId;
        string ParentOfGroup(string groupId) => groupsById[groupId].ParentId ?? RootId;

        string? ImmediateItem(string nodeId, string containerId)
        {
            var parent = ParentOfNode(nodeId);
            if (parent == containerId)
                return nodeId;

            var currentGroup = parent;
            while (currentGroup != RootId && ParentOfGroup(currentGroup) != containerId)
                currentGroup = ParentOfGroup(currentGroup);
            return currentGroup == RootId ? null : currentGroup;
        }

        Box BuildGroup(string groupId, string direction, int order)
        {
            var box = new Box { Id = groupId, IsGroup = groupId != RootId, Order = order };

            foreach (var node in nodes.Where(n => (n.ParentId ?? RootId) == groupId).OrderBy(n => n.Order))
            {
                box.Children.Add(new Box
                {
                    Id = node.Id,
                    IsGroup = false,
                    Order = node.Order,
                    Width = Math.Max(node.Width, 0.1),
                    Height = Math.Max(node.Height, 0.1)
                });
            }

            if (childrenByParent.TryGetValue(groupId, out var childGroups))
            {
                foreach (var child in childGroups)
                    box.Children.Add(BuildGroup(child.Id, child.Direction, child.Order));
            }

            var itemIds = box.Children.Select(c => c.Id).ToHashSet(StringComparer.Ordinal);
            var mappedEdges = new HashSet<(string From, string To)>();
            foreach (var edge in edges)
            {
                var from = ImmediateItem(edge.FromId, groupId);
                var to = ImmediateItem(edge.ToId, groupId);
                if (from != null && to != null && from != to && itemIds.Contains(from) && itemIds.Contains(to))
                    mappedEdges.Add((from, to));
            }

            PlaceItems(box.Children, mappedEdges, direction);
            SizeContainer(box);
            return box;
        }

        var root = BuildGroup(RootId, rootDirection, -1);
        var result = new CompoundGraphLayoutResult();

        void Flatten(Box box, double offsetX, double offsetY)
        {
            var absoluteX = offsetX + box.X;
            var absoluteY = offsetY + box.Y;
            if (box.Id != RootId)
            {
                var rect = new LayoutRect(absoluteX, absoluteY, box.Width, box.Height);
                if (box.IsGroup) result.Groups[box.Id] = rect;
                else result.Nodes[box.Id] = rect;
            }

            foreach (var child in box.Children)
                Flatten(child, absoluteX, absoluteY);
        }

        Flatten(root, 0, 0);
        return result;
    }

    private static void PlaceItems(
        List<Box> items,
        IReadOnlyCollection<(string From, string To)> edges,
        string direction)
    {
        if (items.Count == 0) return;

        var ranks = CalculateRanks(items.Select(i => i.Id), edges);
        var rankGroups = items
            .GroupBy(i => ranks.GetValueOrDefault(i.Id))
            .OrderBy(g => g.Key)
            .Select(g => g.OrderBy(i => i.Order).ToList())
            .ToList();

        var vertical = direction is "TB" or "TD" or "BT";
        var positive = direction is "LR" or "BT";
        var mainCursor = 0.0;

        foreach (var rank in rankGroups)
        {
            var rankMain = rank.Max(i => vertical ? i.Height : i.Width);
            var crossTotal = rank.Sum(i => vertical ? i.Width : i.Height) + ItemGap * (rank.Count - 1);
            var crossCursor = -crossTotal / 2;
            var mainCenter = (mainCursor + rankMain / 2) * (positive ? 1 : -1);

            foreach (var item in rank)
            {
                var crossSize = vertical ? item.Width : item.Height;
                var crossCenter = crossCursor + crossSize / 2;
                if (vertical)
                {
                    item.X = crossCenter;
                    item.Y = mainCenter;
                }
                else
                {
                    item.X = mainCenter;
                    item.Y = crossCenter;
                }
                crossCursor += crossSize + ItemGap;
            }

            mainCursor += rankMain + RankGap;
        }

        // Remove any directional bias from the container's local coordinate system.
        var left = items.Min(i => i.X - i.Width / 2);
        var right = items.Max(i => i.X + i.Width / 2);
        var bottom = items.Min(i => i.Y - i.Height / 2);
        var top = items.Max(i => i.Y + i.Height / 2);
        var centerX = (left + right) / 2;
        var centerY = (bottom + top) / 2;
        foreach (var item in items)
        {
            item.X -= centerX;
            item.Y -= centerY;
        }
    }

    private static void SizeContainer(Box box)
    {
        if (box.Children.Count == 0)
        {
            box.Width = box.IsGroup ? GroupPadding * 2 : 0;
            box.Height = box.IsGroup ? GroupPadding * 2 + GroupTitleHeight : 0;
            return;
        }

        var left = box.Children.Min(i => i.X - i.Width / 2);
        var right = box.Children.Max(i => i.X + i.Width / 2);
        var bottom = box.Children.Min(i => i.Y - i.Height / 2);
        var top = box.Children.Max(i => i.Y + i.Height / 2);

        if (box.IsGroup)
        {
            box.Width = right - left + GroupPadding * 2;
            box.Height = top - bottom + GroupPadding * 2 + GroupTitleHeight;
            var contentCenterY = (bottom + top) / 2 - GroupTitleHeight / 2;
            foreach (var child in box.Children)
                child.Y -= contentCenterY;
        }
        else
        {
            box.Width = right - left;
            box.Height = top - bottom;
        }
    }

    private static Dictionary<string, int> CalculateRanks(
        IEnumerable<string> itemIds,
        IReadOnlyCollection<(string From, string To)> edges)
    {
        var ids = itemIds.ToList();
        var adjacency = ids.ToDictionary(id => id, _ => new List<string>(), StringComparer.Ordinal);
        foreach (var (from, to) in edges)
            if (adjacency.ContainsKey(from) && adjacency.ContainsKey(to)) adjacency[from].Add(to);

        var index = 0;
        var indices = new Dictionary<string, int>(StringComparer.Ordinal);
        var lowLinks = new Dictionary<string, int>(StringComparer.Ordinal);
        var stack = new Stack<string>();
        var onStack = new HashSet<string>(StringComparer.Ordinal);
        var components = new List<List<string>>();

        void StrongConnect(string id)
        {
            indices[id] = index;
            lowLinks[id] = index++;
            stack.Push(id);
            onStack.Add(id);

            foreach (var next in adjacency[id])
            {
                if (!indices.ContainsKey(next))
                {
                    StrongConnect(next);
                    lowLinks[id] = Math.Min(lowLinks[id], lowLinks[next]);
                }
                else if (onStack.Contains(next))
                {
                    lowLinks[id] = Math.Min(lowLinks[id], indices[next]);
                }
            }

            if (lowLinks[id] != indices[id]) return;
            var component = new List<string>();
            string current;
            do
            {
                current = stack.Pop();
                onStack.Remove(current);
                component.Add(current);
            } while (current != id);
            components.Add(component);
        }

        foreach (var id in ids)
            if (!indices.ContainsKey(id)) StrongConnect(id);

        var componentOf = new Dictionary<string, int>(StringComparer.Ordinal);
        for (var i = 0; i < components.Count; i++)
            foreach (var id in components[i]) componentOf[id] = i;

        var predecessors = Enumerable.Range(0, components.Count)
            .ToDictionary(i => i, _ => new HashSet<int>());
        foreach (var (from, to) in edges)
        {
            if (!componentOf.TryGetValue(from, out var fromComponent) ||
                !componentOf.TryGetValue(to, out var toComponent) || fromComponent == toComponent) continue;
            predecessors[toComponent].Add(fromComponent);
        }

        var componentRanks = new Dictionary<int, int>();
        int Rank(int component)
        {
            if (componentRanks.TryGetValue(component, out var rank)) return rank;
            rank = predecessors[component].Count == 0 ? 0 : predecessors[component].Max(Rank) + 1;
            componentRanks[component] = rank;
            return rank;
        }

        return ids.ToDictionary(id => id, id => Rank(componentOf[id]), StringComparer.Ordinal);
    }
}
