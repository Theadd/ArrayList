# [**ArrayList**](https://github.com/Theadd/ArrayList)

<a href="https://github.com/Theadd/ArrayList">
  <img height="28em" align="center" src="https://img.shields.io/badge/GitHub-22272E?style=for-the-badge&logo=&logoColor=white" alt="Github Repository Badge" />
</a>
<a href="https://github.com/Theadd/ArrayList/issues">
  <img height="28em" align="center" src="https://img.shields.io/badge/ISSUES-22272E?style=for-the-badge&logo=&logoColor=white" alt="Issues Badge" />
</a>
<a href="https://github.com/Theadd/ArrayList/releases/latest">
  <img height="28em" align="center" src="https://img.shields.io/badge/RELEASES-22272E?style=for-the-badge&logo=&logoColor=white" alt="Releases Badge" />
</a>
<a href="https://github.com/Theadd/ArrayList/blob/main/LICENSE">
  <img height="28em" align="center" src="https://img.shields.io/badge/UNLICENSE-22272E?style=for-the-badge&logo=&logoColor=white" alt="Unlicense Badge" />
</a>

<br/>
<br/>

`VBA` / `VB6` mscorlib's **ArrayList drop-in replacement** with proper memory management and orders of magnitude faster than `mscorlib.ArrayList`, making use of [twinBASIC](https://github.com/twinbasic/twinbasic)'s new language features and memory management techniques from [VBA-MemoryTools](https://github.com/cristianbuse/VBA-MemoryTools).

Initially, this drop-in replacement for `mscorlib.ArrayList` was just to get rid of the `423 MB VMem` overheat added to VBA projects when using its `ArrayList` implementation. But it also turned out to exceed the speed performance of `mscorlib.ArrayList` by far, along with a proper memory release/deallocation when destroyed or it goes out of scope. Which can't even be manually achieved with `mscorlib.ArrayList` as setting it to `Nothing` or `.Clear` 'ing it doesn't free any memory.


## **Features / Improvements** 

* Takes less than `0.35 MB` of `VMem` to load on first use instead of the `423 MB` taken by `mscorlib`. VBA apps in **Win32** are limited to `2 GB`, if you also add the non-existing memory deallocation of `mscorlib.ArrayList`, continued operations on mid to large datasets are a dead-end in using `mscorlib.ArrayList`.

* As a __drop-in replacement__, it is expected to provide the exact same output and functionality as when using `mscorlib.ArrayList` whithin `VBA`. Static members such as `.Adapter` or `.Repeat` can't be used from `VBA` so they're not included, nor the `Type` parameter in `.ToArray()`, which can't be used either. Additionally, all other members that can could be called or accessed even though they are totally useless from the `VBA` side, are included but hidden, as in duplicated members with similar names to overcome the missing method overloading feature in `COM`, such as `.Sort_2`.

* Unlike `mscorlib.ArrayList`, it allows plain `VBA Arrays` and other enumerable objects as input in parameters expecting a collection-like object _(Of `ICollection` Type in `mscorlib`'s [ArrayList.cs](https://referencesource.microsoft.com/#mscorlib/system/collections/arraylist.cs,215))_.
  ```vb
  ' Example:
  .AddRange(Array( _
      "String at index 0", _
      Array(34, "Lorem Ipsum"), _
      Array(Now(), "Hello World!"), _
      256))
  ```

* Provides an advanced `Enumerator` allowing the use of `For Each` within subranges, backwards enumeration, custom iteration steps and direct access to the backing enumerator instance allowing an even wider set of possibilities while iterating the `Enumerator`.

* The [`Enumerator`](https://github.com/Theadd/ArrayList/blob/main/ArrayListLib/Sources/Enumerator.twin#L21) class is publicly accessible so you can reuse it anywhere else in your code.

* Using multidimensional arrays as elements is <u>not</u> supported by `mscorlib.ArrayList` but `ArrayList` seems to have no reason for that, they just work like any other value or reference types when being added as elements. <sup><small>(If anyone encounters with such problems in `ArrayList` please post an issue)</small></sup>

* The `.GetRange` method from `mscorlib.ArrayList` returns a [`Range`]() class instance, which extends the `mscorlib.ArrayList` class so the return of `.GetRange` can be directly assigned to a variable declared as `mscorlib.ArrayList`. In order to achieve this in our `ArrayList` class, it has been declared as a `CoClass`, which gives no-intellisense support, but is expected to be fixed soon.


## **Documentation**

* `ArrayList` docs are available [here](Docs/ArrayList.md) but, as a drop-in replacement, you can also use the ones from the official [`.NET Documentation`](https://learn.microsoft.com/en-us/dotnet/api/system.collections.arraylist?view=netframework-4.7.2) which has lots of usage examples, just ignore `static` members as they can't be called from `VBA`. You might also like to have a look directly at the [source code](https://referencesource.microsoft.com/#mscorlib/system/collections/arraylist.cs) of `mscorlib.ArrayList` in **CSharp**.  


## **Performance Tests Results**

**Percentages are calculated against `mscorlib.ArrayList`'s timings from corresponding _Win64_ or _Win32_ results.**

The lower the percentage, the better. **20%** equals to a 5 times faster performance while **500%** would be 5 times slower than their corresponding _Win64_ or _Win32_ execution time in `mscorlib.ArrayList`.


<table>
    <thead>
        <tr>
            <th>&nbsp;</th>
            <th colspan=4>ArrayList Class</th>
            <th colspan=2>mscorlib.ArrayList</th>
            <th colspan=4>VBA.Collection</th>
        </tr>
        <tr>
            <th>&nbsp;</th>
            <th colspan=2>x64</th>
            <th colspan=2>x32</th>
            <th>x64</th>
            <th>x32</th>
            <th colspan=2>x64</th>
            <th colspan=2>x32</th>
        </tr>
    </thead>
    <tbody>
        <tr>
            <td><code>.Add</code> <sup><small><br/>(x5000)</small></sup></td>
            <td>0&nbsp;ms</td>
            <td><b>6%</b></td>
            <td>0&nbsp;ms</td>
            <td><b>4%</b></td>
            <td>17&nbsp;ms</td>
            <td>23&nbsp;ms</td>
            <td>0&nbsp;ms</td>
            <td><b>6%</b></td>
            <td>1&nbsp;ms</td>
            <td><b>4%</b></td>
        </tr>
        <tr>
            <td><code>.Add</code> <sup><small><br/>(250x5000)</small></sup></td>
            <td>134&nbsp;ms</td>
            <td><b>4%</b></td>
            <td>181&nbsp;ms</td>
            <td><b>3%</b></td>
            <td>3579&nbsp;ms</td>
            <td>6302&nbsp;ms</td>
            <td>186&nbsp;ms</td>
            <td><b>5%</b></td>
            <td>184&nbsp;ms</td>
            <td><b>3%</b></td>
        </tr>
        <tr>
            <td><code>.Clone</code> <sup><small><br/>(50x5000)</small></sup></td>
            <td>3&nbsp;ms</td>
            <td><b>100%</b></td>
            <td>5&nbsp;ms</td>
            <td><b>167%</b></td>
            <td>3&nbsp;ms</td>
            <td>3&nbsp;ms</td>
            <td>2773&nbsp;ms</td>
            <td><b>92433%</b></td>
            <td>2674&nbsp;ms</td>
            <td><b>89133%</b></td>
        </tr>
        <tr>
            <td><code>.Insert&nbsp;Index:=0</code> <sup><small><br/>(20x1000)</small></sup></td>
            <td>9&nbsp;ms</td>
            <td><b>4%</b></td>
            <td>10&nbsp;ms</td>
            <td><b>6%</b></td>
            <td>209&nbsp;ms</td>
            <td>181&nbsp;ms</td>
            <td>3&nbsp;ms</td>
            <td><b>1%</b></td>
            <td>4&nbsp;ms</td>
            <td><b>2%</b></td>
        </tr>
        <tr>
            <td><code>.Insert&nbsp;Index:=RND</code> <sup><small><br/>(2x5000)</small></sup></td>
            <td>37&nbsp;ms</td>
            <td><b>49%</b></td>
            <td>43&nbsp;ms</td>
            <td><b>41%</b></td>
            <td>75&nbsp;ms</td>
            <td>104&nbsp;ms</td>
            <td>147&nbsp;ms</td>
            <td><b>196%</b></td>
            <td>124&nbsp;ms</td>
            <td><b>119%</b></td>
        </tr>
        <tr>
            <td><code>.Item</code> <sup><small><br/>(Read) SEQ (20x5000)</small></sup></td>
            <td>9&nbsp;ms</td>
            <td><b>2%</b></td>
            <td>10&nbsp;ms</td>
            <td><b>2%</b></td>
            <td>421&nbsp;ms</td>
            <td>644&nbsp;ms</td>
            <td>1125&nbsp;ms</td>
            <td><b>267%</b></td>
            <td>1052&nbsp;ms</td>
            <td><b>163%</b></td>
        </tr>
        <tr>
            <td><code>.Item</code> <sup><small><br/>(Read) SEQ+RND (20x5000)</small></sup></td>
            <td>19&nbsp;ms</td>
            <td><b>3%</b></td>
            <td>21&nbsp;ms</td>
            <td><b>2%</b></td>
            <td>735&nbsp;ms</td>
            <td>1135&nbsp;ms</td>
            <td>2198&nbsp;ms</td>
            <td><b>299%</b></td>
            <td>2142&nbsp;ms</td>
            <td><b>189%</b></td>
        </tr>
        <tr>
            <td><code>.RemoveAt&nbsp;Index:=0</code> <sup><small><br/>(x5000)</small></sup></td>
            <td>5&nbsp;ms</td>
            <td><b>22%</b></td>
            <td>4&nbsp;ms</td>
            <td><b>13%</b></td>
            <td>23&nbsp;ms</td>
            <td>31&nbsp;ms</td>
            <td>1&nbsp;ms</td>
            <td><b>4%</b></td>
            <td>1&nbsp;ms</td>
            <td><b>3%</b></td>
        </tr>
        <tr>
            <td><code>.RemoveAt&nbsp;Index:=RND</code> <sup><small><br/>(x12000)</small></sup></td>
            <td>70&nbsp;ms</td>
            <td><b>34%</b></td>
            <td>42&nbsp;ms</td>
            <td><b>20%</b></td>
            <td>203&nbsp;ms</td>
            <td>209&nbsp;ms</td>
            <td>910&nbsp;ms</td>
            <td><b>448%</b></td>
            <td>675&nbsp;ms</td>
            <td><b>323%</b></td>
        </tr>
        <tr>
            <td><code>.RemoveAt&nbsp;Index:=LAST</code> <sup><small><br/>(x12000)</small></sup></td>
            <td>1&nbsp;ms</td>
            <td><b>3%</b></td>
            <td>2&nbsp;ms</td>
            <td><b>3%</b></td>
            <td>33&nbsp;ms</td>
            <td>58&nbsp;ms</td>
            <td>508&nbsp;ms</td>
            <td><b>1539%</b></td>
            <td>340&nbsp;ms</td>
            <td><b>586%</b></td>
        </tr>
        <tr>
            <td><code>.GetRange&nbsp;Index:=RND</code> <sup><small><br/>(1000x5000)</small></sup></td>
            <td>2&nbsp;ms</td>
            <td><b>18%</b></td>
            <td>3&nbsp;ms</td>
            <td><b>16%</b></td>
            <td>11&nbsp;ms</td>
            <td>19&nbsp;ms</td>
            <td></td>
            <td></td>
            <td></td>
            <td></td>
        </tr>
        <tr>
            <td><code>.AddRange&nbsp;Range</code> <sup><small><br/>(x1000)</small></sup></td>
            <td>77&nbsp;ms</td>
            <td><b>31%</b></td>
            <td>176&nbsp;ms</td>
            <td><b>103%</b></td>
            <td>249&nbsp;ms</td>
            <td>171&nbsp;ms</td>
            <td></td>
            <td></td>
            <td></td>
            <td></td>
        </tr>
        <tr>
            <td><code>.GetEnumerator</code> <sup><small><br/>(For Each) (10x100000)</small></sup></td>
            <td>111&nbsp;ms</td>
            <td><b>14%</b></td>
            <td>114&nbsp;ms</td>
            <td><b>17%</b></td>
            <td>775&nbsp;ms</td>
            <td>659&nbsp;ms</td>
            <td>25&nbsp;ms</td>
            <td><b>3%</b></td>
            <td>25&nbsp;ms</td>
            <td><b>4%</b></td>
        </tr>
        <tr>
            <td><code>.Sort</code> <sup><small><br/>(100000)</small></sup></td>
            <td>167&nbsp;ms</td>
            <td><b>293%</b></td>
            <td>181&nbsp;ms</td>
            <td><b>503%</b></td>
            <td>57&nbsp;ms</td>
            <td>36&nbsp;ms</td>
            <td></td>
            <td></td>
            <td></td>
            <td></td>
        </tr>
        <tr>
            <td><code>.Sort&nbsp;w/&nbsp;Comparer</code> <sup><small><br/>(100000)</small></sup></td>
            <td>494&nbsp;ms</td>
            <td><b>37%</b></td>
            <td>453&nbsp;ms</td>
            <td><b>18%</b></td>
            <td>1320&nbsp;ms</td>
            <td>2473&nbsp;ms</td>
        </tr>
    </tbody>
</table>


**Where <kbd> `.Add` <sup><small>(250x5000)</small></sup></kbd> equals to the following code.**

```vb
For e = 0 To Iterations - 1     ' Iterations = 250
    For i = 0 To UBound(t)      ' UBound(t) = 5000 - 1
        .Add t(i)
    Next i
Next e
```
*That's a total of **1,250,000** calls to `.Add`, taking only **134ms** (Win64) instead of **3,579ms** in `mscorlib.ArrayList`.*


While the `.Add` method of `VBA.Collection` has similar performance as in our `ArrayList`, reading/accessing their values is potentially slow and gets exponentially worse depending on the number of elements it contains. A simple sequential read of **5,000** items in a `VBA.Collection` takes **59 ms**, reading **10,000** items takes **266 ms**, which is almost 5 times more with just twice the size but reading **100,000** items, takes **36,069 ms**, 135 times slower just increasing 10 times it's size. Our `ArrayList` only takes **9 ms** to read **100,000** items, increasing linearly, taking **94 ms** to read **1,000,000** items.

So in order to include `VBA.Collection` in the table above, tests are iterated multiple times over **5,000** items instead of using bigger sizes.


## **Acknowledgments**

* To [@CristianBuse](https://github.com/cristianbuse)'s [`VBA-MemoryTools`](https://github.com/cristianbuse/VBA-MemoryTools), from which I discovered a whole new level in `VBA` programming, is what runs the most performance-critical parts behind `ArrayList`, and himself for his amazing and extensive support.


## **License**

- [`LibMemory`](Sources/LibMemory.twin) from [VBA-MemoryTools](https://github.com/cristianbuse/VBA-MemoryTools) is released under the [MIT License](https://github.com/cristianbuse/VBA-MemoryTools/blob/master/LICENSE).

- Everything else is released under [The Unlicense](https://github.com/Theadd/ArrayList/blob/main/LICENSE) into the public domain.
