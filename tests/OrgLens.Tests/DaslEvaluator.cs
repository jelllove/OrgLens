using System;
using System.Collections.Generic;
using System.Text;

namespace OrgLens.Tests
{
    internal enum SqlTruth { False, True, Unknown }

    // This deliberately small parser evaluates only the generated DASL grammar, not C# or script.
    internal sealed class DaslExpression
    {
        private readonly Func<IReadOnlyDictionary<string, object>, SqlTruth> evaluate;

        private DaslExpression(Func<IReadOnlyDictionary<string, object>, SqlTruth> evaluate,
            IReadOnlyCollection<string> properties)
        {
            this.evaluate = evaluate;
            Properties = properties;
        }

        internal IReadOnlyCollection<string> Properties { get; }
        internal SqlTruth Evaluate(IReadOnlyDictionary<string, object> message) { return evaluate(message); }
        internal bool Matches(IReadOnlyDictionary<string, object> message)
        {
            return Evaluate(message) == SqlTruth.True;
        }
        internal static DaslExpression Parse(string filter) { return new Parser(filter).Parse(); }

        private static SqlTruth Not(SqlTruth value)
        {
            return value == SqlTruth.Unknown ? SqlTruth.Unknown :
                value == SqlTruth.True ? SqlTruth.False : SqlTruth.True;
        }

        private static SqlTruth And(SqlTruth left, SqlTruth right)
        {
            if (left == SqlTruth.False || right == SqlTruth.False) return SqlTruth.False;
            return left == SqlTruth.Unknown || right == SqlTruth.Unknown ? SqlTruth.Unknown : SqlTruth.True;
        }

        private static SqlTruth Or(SqlTruth left, SqlTruth right)
        {
            if (left == SqlTruth.True || right == SqlTruth.True) return SqlTruth.True;
            return left == SqlTruth.Unknown || right == SqlTruth.Unknown ? SqlTruth.Unknown : SqlTruth.False;
        }

        private sealed class Parser
        {
            private readonly string input;
            private readonly HashSet<string> properties = new HashSet<string>(StringComparer.Ordinal);
            private int position;

            internal Parser(string input)
            {
                this.input = input ?? throw new ArgumentNullException(nameof(input));
            }

            internal DaslExpression Parse()
            {
                var expression = ParseOr();
                SkipWhitespace();
                if (position != input.Length) throw Error("Unexpected trailing token");
                return new DaslExpression(expression, new List<string>(properties));
            }

            private Func<IReadOnlyDictionary<string, object>, SqlTruth> ParseOr()
            {
                var expression = ParseAnd();
                while (TakeWord("OR"))
                {
                    var left = expression;
                    var right = ParseAnd();
                    expression = message => Or(left(message), right(message));
                }
                return expression;
            }

            private Func<IReadOnlyDictionary<string, object>, SqlTruth> ParseAnd()
            {
                var expression = ParseUnary();
                while (TakeWord("AND"))
                {
                    var left = expression;
                    var right = ParseUnary();
                    expression = message => And(left(message), right(message));
                }
                return expression;
            }

            private Func<IReadOnlyDictionary<string, object>, SqlTruth> ParseUnary()
            {
                if (TakeWord("NOT"))
                {
                    var operand = ParseUnary();
                    return message => Not(operand(message));
                }
                if (Take('('))
                {
                    var expression = ParseOr();
                    Require(')');
                    return expression;
                }
                return ParseComparison();
            }

            private Func<IReadOnlyDictionary<string, object>, SqlTruth> ParseComparison()
            {
                string property = ReadQuoted('"');
                properties.Add(property);
                if (TakeWord("IS"))
                {
                    bool not = TakeWord("NOT");
                    if (!TakeWord("NULL")) throw Error("Expected NULL");
                    return message =>
                    {
                        object actual;
                        bool isNull = !message.TryGetValue(property, out actual) || actual == null;
                        return isNull != not ? SqlTruth.True : SqlTruth.False;
                    };
                }
                Require('=');
                SkipWhitespace();
                object expected;
                if (position < input.Length && input[position] == '\'') expected = ReadQuoted('\'');
                else if (Take('0')) expected = 0;
                else if (Take('1')) expected = 1;
                else throw Error("Expected a quoted string, 0, or 1");
                return message =>
                {
                    object actual;
                    if (!message.TryGetValue(property, out actual) || actual == null) return SqlTruth.Unknown;
                    bool equal = actual is string && expected is string
                        ? string.Equals((string)actual, (string)expected, StringComparison.OrdinalIgnoreCase)
                        : Equals(actual, expected);
                    return equal ? SqlTruth.True : SqlTruth.False;
                };
            }

            private string ReadQuoted(char quote)
            {
                Require(quote);
                var value = new StringBuilder();
                while (position < input.Length)
                {
                    char next = input[position++];
                    if (next != quote) value.Append(next);
                    else if (position < input.Length && input[position] == quote)
                    {
                        position++;
                        value.Append(quote);
                    }
                    else return value.ToString();
                }
                throw Error("Unterminated quoted value");
            }

            private bool TakeWord(string word)
            {
                SkipWhitespace();
                if (position + word.Length > input.Length ||
                    string.Compare(input, position, word, 0, word.Length, StringComparison.OrdinalIgnoreCase) != 0)
                    return false;
                int end = position + word.Length;
                if (end < input.Length && (char.IsLetterOrDigit(input[end]) || input[end] == '_')) return false;
                position = end;
                return true;
            }

            private bool Take(char character)
            {
                SkipWhitespace();
                if (position == input.Length || input[position] != character) return false;
                position++;
                return true;
            }

            private void Require(char character)
            {
                if (!Take(character)) throw Error("Expected " + character);
            }

            private void SkipWhitespace()
            {
                while (position < input.Length && char.IsWhiteSpace(input[position])) position++;
            }

            private FormatException Error(string message)
            {
                return new FormatException(message + " at DASL offset " + position + ".");
            }
        }
    }
}
