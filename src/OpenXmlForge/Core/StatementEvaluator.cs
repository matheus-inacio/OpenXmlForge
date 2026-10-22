using System.Globalization;

namespace OpenXmlForge.Core;

internal static class StatementEvaluator
{
    private const string InlineIfStart = "{{if(";
    private const string InlineIfMid = ")if";
    private const string InlineIfEnd = "endif}}";

    public static bool EvaluateComparison(string? leftRaw, string comparisonOperator, string? rightRaw)
    {
        var left = leftRaw ?? string.Empty;
        var right = rightRaw ?? string.Empty;
        var op = comparisonOperator.Trim();

        // Numeric comparison
        if (TryParseDouble(left, out var leftNum) && TryParseDouble(right, out var rightNum))
        {
            return op switch
            {
                "==" => leftNum.Equals(rightNum),
                "!=" => !leftNum.Equals(rightNum),
                ">" => leftNum > rightNum,
                "<" => leftNum < rightNum,
                ">=" => leftNum >= rightNum,
                "<=" => leftNum <= rightNum,
                _ => false,
            };
        }

        // DateTime comparison
        if (TryParseDate(left, out var leftDate) && TryParseDate(right, out var rightDate))
        {
            return op switch
            {
                "==" => leftDate.Equals(rightDate),
                "!=" => !leftDate.Equals(rightDate),
                ">" => leftDate > rightDate,
                "<" => leftDate < rightDate,
                ">=" => leftDate >= rightDate,
                "<=" => leftDate <= rightDate,
                _ => false,
            };
        }

        // Boolean comparison
        if (bool.TryParse(left, out var leftBool) && bool.TryParse(right, out var rightBool))
        {
            return op switch
            {
                "==" => leftBool == rightBool,
                "!=" => leftBool != rightBool,
                _ => false,
            };
        }

        // String comparison
        return op switch
        {
            "==" => string.Equals(left, right, StringComparison.Ordinal),
            "!=" => !string.Equals(left, right, StringComparison.Ordinal),
            _ => false,
        };
    }

    public static string EvaluateInlineIfStatements(
        string text,
        IReadOnlyDictionary<string, object?>? scope = null)
    {
        if (!text.Contains(InlineIfStart, StringComparison.Ordinal))
        {
            return text;
        }

        while (true)
        {
            var ifIndex = text.IndexOf(InlineIfStart, StringComparison.Ordinal);
            if (ifIndex < 0)
            {
                break;
            }

            var endParenIndex = text.IndexOf(')', ifIndex + InlineIfStart.Length);
            if (endParenIndex < 0)
            {
                break;
            }

            // Support both classic "{{if(a,==,b)if body endif}}" and natural "{{if(a == b) body endif}}"
            var hasClassicIfSuffix =
                endParenIndex + 3 <= text.Length &&
                text.AsSpan(endParenIndex, 3).SequenceEqual(InlineIfMid);
            var bodyStart = hasClassicIfSuffix ? endParenIndex + InlineIfMid.Length : endParenIndex + 1;

            var endIfFinalIndex = text.IndexOf(InlineIfEnd, bodyStart, StringComparison.Ordinal);
            if (endIfFinalIndex < 0)
            {
                break;
            }

            var conditionSpan = text.Substring(
                ifIndex + InlineIfStart.Length,
                endParenIndex - (ifIndex + InlineIfStart.Length));

            var parts = conditionSpan.Contains(',')
                ? conditionSpan.Split(',')
                : conditionSpan.Split(' ', StringSplitOptions.RemoveEmptyEntries);

            var isTrue = false;
            if (parts.Length >= 3)
            {
                var leftOperand = ResolveOperandIfNeeded(parts[0].Trim(), scope);
                var rightOperand = ResolveOperandIfNeeded(parts[2].Trim(), scope);
                isTrue = EvaluateComparison(leftOperand, parts[1].Trim(), rightOperand);
            }

            if (isTrue)
            {
                var bodyContent = text[bodyStart..endIfFinalIndex];
                if (!hasClassicIfSuffix)
                {
                    bodyContent = bodyContent.Trim();
                }

                text = string.Concat(
                    text.AsSpan(0, ifIndex),
                    bodyContent,
                    text.AsSpan(endIfFinalIndex + InlineIfEnd.Length));
            }
            else
            {
                text = text.Remove(ifIndex, endIfFinalIndex + InlineIfEnd.Length - ifIndex);
            }
        }

        return text;
    }

    private static string ResolveOperandIfNeeded(
        string operand,
        IReadOnlyDictionary<string, object?>? scope)
    {
        if (scope is not null &&
            !bool.TryParse(operand, out _) &&
            !TryParseDouble(operand, out _) &&
            ValueResolver.TryResolveValue(scope, operand, out var resolved) &&
            resolved is not null)
        {
            return Convert.ToString(resolved, CultureInfo.InvariantCulture) ?? string.Empty;
        }

        return operand;
    }

    private static bool TryParseDouble(string text, out double number)
        => double.TryParse(text, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out number);

    private static bool TryParseDate(string text, out DateTime date)
        => DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out date) ||
           DateTime.TryParse(text, CultureInfo.CurrentCulture, DateTimeStyles.None, out date);
}
