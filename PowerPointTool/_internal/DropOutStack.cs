using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace PowerPointTool._internal;

class DropOutStack<T>(int capacity) : IReadOnlyCollection<T>
{
    readonly T[] _items = new T[capacity];
    int _top = 0;
    int _count = 0;

    public void Push(T item)
    {
        _items[_top] = item;
        _top = (_top + 1) % capacity;

        if (_count < capacity)
            _count++;
    }

    public T Pop()
    {
        if (_count == 0)
            return default;

        _top = (capacity + _top - 1) % capacity;
        _count--;

        return _items[_top];
    }

    public T Peek(int i = 0)
    {
        if (i >= capacity)
            return default;

        return _items[(capacity + _top - 1 - i) % capacity];
    }

    public int Capacity => capacity;

    public int Count => _count;

    public T this[int index] => Peek(index);

    public IEnumerator<T> GetEnumerator() => Enumerable.Range(0, _count).Select(Peek).GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    public override string ToString() => string.Join(", ", this);
}