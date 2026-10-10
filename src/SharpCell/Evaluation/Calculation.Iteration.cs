using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;

namespace SharpCell.Evaluation;

/// <summary>
/// Iterative calculation (<see cref="Workbook.Iteration"/>): a circular reference is calculated in
/// passes, each cell reading the others' latest values (Gauss–Seidel), until no value changes by
/// more than <see cref="IterationSettings.MaxChange"/> or <see cref="IterationSettings.MaxIterations"/>
/// passes are done. Its cells are then volatile: each recalculation continues from their values.
/// </summary>
internal sealed partial class Calculation
{
    /// <summary>The cycle being iterated: reading one of its cells gives its value instead of waiting for it.</summary>
    public HashSet<CellKey>? ActiveCycle { get; private set; }

    /// <summary>Thrown when a cell turns out to belong to the cycle being iterated; the iteration starts again with it.</summary>
    private sealed class CycleGrew(CellKey joiner) : Exception("The cycle being iterated has more cells.")
    {
        public CellKey Joiner { get; } = joiner;
    }

    private void IterateCycle(CellKey asker, CellKey asked, CancellationToken cancellationToken)
    {
        var members = new HashSet<CellKey>(CycleMembers(asker, asked));
        while (true)
        {
            try
            {
                if (!Iterate(members, cancellationToken))
                    ResolveCycle(asker, asked);
                return;
            }
            catch (CycleGrew grew)
            {
                if (!members.Add(grew.Joiner))
                    throw new InvalidOperationException($"{grew.Joiner} joined the cycle twice.", grew);
            }
        }
    }

    // Returns false when the cycle cannot be iterated: a member's result spills.
    private bool Iterate(HashSet<CellKey> members, CancellationToken cancellationToken)
    {
        var settings = workbook.Iteration;
        var order = new List<CellKey>(members);
        var sheets = SheetOrder();
        order.Sort((a, b) =>
        {
            var bySheet = sheets[a.Sheet].CompareTo(sheets[b.Sheet]);
            return bySheet != 0 ? bySheet : a.Row != b.Row ? a.Row.CompareTo(b.Row) : a.Column.CompareTo(b.Column);
        });

        var results = new Dictionary<CellKey, (CellValue Value, Dependencies Dependencies, IReadOnlyList<CalculationDiagnostic> Diagnostics)>();
        foreach (var member in order)
            Remember(member, member.Data);

        ActiveCycle = members;
        try
        {
            var passes = 0;
            double change;
            do
            {
                cancellationToken.ThrowIfCancellationRequested();
                passes++;
                change = 0;
                foreach (var member in order)
                {
                    var data = member.Data!;
                    var (value, context) = EvaluateMember(member, data, cancellationToken);
                    if (value.Kind == CellValueKind.Array && data.FixedArray is null)
                        return false;

                    var scalar = value.Kind == CellValueKind.Array ? value.AsArray()[0, 0] : value;
                    change = Math.Max(change, Change(data.Value, scalar));
                    data.Value = scalar;
                    results[member] = (value, context.Dependencies!, context.Diagnostics);
                }
            }
            while (change > settings.MaxChange && passes < settings.MaxIterations);

            ActiveCycle = null;
            var limitReached = change > settings.MaxChange;
            var diagnostic = limitReached ? LimitDiagnostic(order, settings) : null;
            var quiet = new HashSet<CellKey>(members);
            foreach (var member in order)
            {
                var (value, dependencies, diagnostics) = results[member];
                Commit(member, member.Data!, value, dependencies, usedVolatile: true, cancellationToken, quiet);
                SetDiagnostics(member, diagnostic is null
                    ? diagnostics
                    : [new CalculationDiagnostic(DiagnosticKind.IterationLimitReached, member.Sheet, member.Address.ToString(), diagnostic)]);
            }

            return true;
        }
        finally
        {
            ActiveCycle = null;
        }
    }

    // Evaluates one member, first calculating the cells outside the cycle it reads that are out of
    // date; one of those that reads the cycle joins it.
    private (CellValue Value, EvaluationContext Context) EvaluateMember(CellKey member, CellData data, CancellationToken cancellationToken)
    {
        while (true)
        {
            var context = new EvaluationContext(workbook, member.Sheet, member.Address)
            {
                CancellationToken = cancellationToken,
                Dependencies = new Dependencies(),
                Legacy = data.IsLegacy,
                LegacyFormula = data.IsLegacy,
                IsCycleMember = true,
            };
            EvaluationCount++;
            var value = EvaluateGuarded(data.Formula!, context);
            if (context.Pending.Count == 0)
                return (value, context);
            foreach (var input in context.Pending)
                Compute(input, cancellationToken);
        }
    }

    // Numbers by their difference, anything else settled only when unchanged. An empty cell starts as 0.
    private static double Change(CellValue before, CellValue after)
    {
        if (before.Kind == CellValueKind.Empty)
            before = CellValue.Number(0);
        if (after.Kind == CellValueKind.Empty)
            after = CellValue.Number(0);
        if (before.Kind == CellValueKind.Number && after.Kind == CellValueKind.Number)
            return Math.Abs(after.AsNumber() - before.AsNumber());
        return before.Equals(after) ? 0 : double.PositiveInfinity;
    }

    private static string LimitDiagnostic(List<CellKey> order, IterationSettings settings)
    {
        var text = new StringBuilder($"Circular reference did not settle in {settings.MaxIterations} iterations: ");
        for (var i = 0; i < order.Count; i++)
            text.Append(i == 0 ? "" : ", ").Append(order[i]);
        return text.ToString();
    }
}
