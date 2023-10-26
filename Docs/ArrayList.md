# **`ArrayList` Class**

_Backwards-compatible custom version of [Tim Hall](https://github.com/VBA-tools/VBA-JSON)'s `JSON` utilities._

---

## **API Overview**

```vb
' Properties
Property Get Item(ByVal Index As Long) As Variant   ' Get, Let & Set
Property Get Count() As Long
Property Get Capacity() As Long                     ' Get & Let
' Methods
Function Add(Value As Variant) As Long
Sub Clear()
Function Clone() As Variant
Function Contains(Value As Variant) As Boolean
Sub CopyTo(Target As Variant, Index As Long)
Function GetEnumerator(Optional ByVal Index As Long = 0, Optional ByVal GetCount As Variant, Optional ByVal GetStep As Long = 1, Optional ByRef ThisEnumerator As IEnumerator) As stdole.IUnknown
Function IndexOf(Value As Variant, Optional ByVal Index As Long = 0, Optional ByVal GetCount As Variant) As Long
Function LastIndexOf(ByRef Value As Variant, Optional ByVal Index As Variant, Optional ByVal GetCount As Variant) As Long
Sub Insert(ByVal Index As Long, Value As Variant)
Sub Remove(Value As Variant)
Sub RemoveAt(ByVal Index As Long)
Sub RemoveRange(ByVal Index As Long, ByVal GetCount As Long)
Sub Reverse(Optional ByVal Index As Long = 0, Optional ByVal GetCount As Variant)
Function ToArray() As Variant()
Sub AddRange(Target As Variant)
Sub InsertRange(ByVal Index As Long, Target As Variant)
Function GetRange(ByVal Index As Long, ByVal GetCount As Long) As IListRange
Sub SetRange(ByVal Index As Long, Target As Variant)
Sub Sort(Optional ByVal Index As Long = 0, Optional ByVal GetCount As Variant, Optional Comparer As IComparer = Nothing)
Function BinarySearch(ByVal Index As Long, ByVal GetCount As Long, Value As Variant, Optional ByRef Comparer As IComparer = Nothing) As Long
```


## **`ArrayList` API**  


### **Properties**

#### Item

```vb
[ DefaultMember ]
Get Item(ByVal Index As Long) As Variant
Let Item(ByVal Index As Long, Value As Variant)
Set Item(ByVal Index As Long, Value As Variant)
```


#### Count

```vb
Get Count() As Long
```



#### Capacity

```vb
Get Capacity() As Long
Let Capacity(Value As Long)
```


#### **Methods**


#### Add

```vb
Function Add(Value As Variant) As Long
```

Adds an item to the list. The return value is the position the new element was inserted in.


#### Clear

```vb
Sub Clear()
```

Removes all items from the list.



#### Clone

```vb
Function Clone() As Variant
```

Creates a shallow copy of this ArrayList.







#### Contains

```vb
Function Contains(Value As Variant) As Boolean
```

Returns whether the list contains a particular item using strict equality comparison.

#### CopyTo

```vb
Sub CopyTo(Target As Variant, Index As Long)
```

Copies this ArrayList to another array at specified index, the other array must be of a compatible array type but not necessarily the same type. It also accepts other lists implementing IListRange as target.


#### GetEnumerator

```vb
[ Enumerator ]
Function GetEnumerator(Optional ByVal Index As Long = 0, _
                       Optional ByVal GetCount As Variant, _
                       Optional ByVal GetStep As Long = 1, _
                       Optional ByRef ThisEnumerator As IEnumerator) As stdole.IUnknown
```

#### IndexOf

```vb
Function IndexOf(Value As Variant, _
                 Optional ByVal Index As Long = 0, _
                 Optional ByVal GetCount As Variant) As Long
```

Returns the index of a particular item or -1 if the item isn't in the list, using strict equality comparison.

#### LastIndexOf

```vb
Function LastIndexOf(ByRef Value As Variant, _
                     Optional ByVal Index As Variant, _
                     Optional ByVal GetCount As Variant) As Long
```

Returns the last index of a particular item or -1 if the item isn't in the list, using strict equality comparison.


#### Insert

```vb
Sub Insert(ByVal Index As Long, Value As Variant)
```

Inserts value into the list at position Index. Index must be non-negative and less than or equal to the number of elements in the list. If Index equals the number of items in the list, then value is appended to the end.


#### Remove

```vb
Sub Remove(Value As Variant)
```

Removes an item from the list.


#### RemoveAt

```vb
Sub RemoveAt(ByVal Index As Long)
```

Removes the item at Index position.


#### RemoveRange

```vb
Sub RemoveRange(ByVal Index As Long, ByVal GetCount As Long)
```




#### Reverse

```vb
Sub Reverse(Optional ByVal Index As Long = 0, Optional ByVal GetCount As Variant)
```




#### ToArray

```vb
Function ToArray() As Variant()
```



#### InsertRange

```vb
Sub InsertRange(ByVal Index As Long, Target As Variant)
```



#### GetRange

```vb
Function GetRange(ByVal Index As Long, ByVal GetCount As Long) As IListRange
```



#### SetRange

```vb
Sub SetRange(ByVal Index As Long, Target As Variant)
```



#### Sort

```vb
Sub Sort(Optional ByVal Index As Long = 0, Optional ByVal GetCount As Variant, Optional Comparer As IComparer = Nothing)
```



#### BinarySearch

```vb
Function BinarySearch(ByVal Index As Long, ByVal GetCount As Long, Value As Variant, Optional ByRef Comparer As IComparer = Nothing) As Long
```

Searches a section of a sorted list. Returns the index of the given value in the list. If not found, returns a negative integer. Use the bitwise operator (`Not`) to get the index of the first element larger than this one, if any.



