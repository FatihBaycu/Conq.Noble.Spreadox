using System.Collections.Generic;

namespace Conq.Noble.Spreadox.Internal
{
    internal sealed class SharedStringTableBuilder
    {
        private readonly Dictionary<string, int> _indexByValue;
        private readonly List<string> _values;

        public SharedStringTableBuilder()
            : this(0)
        {
        }

        public SharedStringTableBuilder(int capacity)
        {
            _indexByValue = capacity > 0
                ? new Dictionary<string, int>(capacity)
                : new Dictionary<string, int>();
            _values = capacity > 0 ? new List<string>(capacity) : new List<string>();
        }

        public int Count => _values.Count;

        public int GetOrAdd(string value)
        {
            if (_indexByValue.TryGetValue(value, out int index))
            {
                return index;
            }

            index = _values.Count;
            _values.Add(value);
            _indexByValue[value] = index;
            return index;
        }

        public IReadOnlyList<string> Values => _values;
    }
}
