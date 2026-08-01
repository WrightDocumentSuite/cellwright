namespace Cellwright.Formulas.Engine;

/// <summary>
/// Topological ordering of formulas with cycle detection. The traversal keeps its own stack,
/// so a chain of a hundred thousand dependent formulas costs heap memory rather than call
/// frames.
/// </summary>
internal static class CalculationOrder
{
    private enum Mark : byte
    {
        White,
        Gray,
        Black,
    }

    /// <summary>
    /// Returns the evaluation order, dependencies first, and the set of cells caught in a
    /// circular reference.
    /// </summary>
    public static (int[] Order, HashSet<int> Circular) Build(IReadOnlyList<FormulaEntry> entries)
    {
        var state = new Mark[entries.Count];
        var order = new List<int>(entries.Count);
        var circular = new HashSet<int>();
        var nodes = new List<int>();
        var cursors = new List<int>();

        for (int root = 0; root < entries.Count; root++)
        {
            if (state[root] != Mark.White)
                continue;

            state[root] = Mark.Gray;
            nodes.Add(root);
            cursors.Add(0);
            Traverse(entries, state, order, circular, nodes, cursors);
        }

        return ([.. order], circular);
    }

    private static void Traverse(
        IReadOnlyList<FormulaEntry> entries, Mark[] state, List<int> order,
        HashSet<int> circular, List<int> nodes, List<int> cursors)
    {
        while (nodes.Count > 0)
        {
            int node = nodes[^1];
            List<int> precedents = entries[node].Precedents;
            int cursor = cursors[^1];

            if (cursor >= precedents.Count)
            {
                state[node] = Mark.Black;
                order.Add(node);
                nodes.RemoveAt(nodes.Count - 1);
                cursors.RemoveAt(cursors.Count - 1);
                continue;
            }

            cursors[^1] = cursor + 1;
            int precedent = precedents[cursor];
            switch (state[precedent])
            {
                case Mark.White:
                    state[precedent] = Mark.Gray;
                    nodes.Add(precedent);
                    cursors.Add(0);
                    break;

                case Mark.Gray:
                    MarkCycle(nodes, precedent, circular);
                    break;

                default:
                    break;
            }
        }
    }

    /// <summary>Marks every cell on the open path from the revisited node onwards.</summary>
    private static void MarkCycle(List<int> nodes, int start, HashSet<int> circular)
    {
        int from = nodes.LastIndexOf(start);
        if (from < 0)
        {
            circular.Add(start);
            return;
        }

        for (int i = from; i < nodes.Count; i++)
            circular.Add(nodes[i]);
    }
}
