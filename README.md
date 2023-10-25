# **ArrayList**
`VBA` / `VB6` mscorlib's **ArrayList drop-in replacement**, implemented using [twinBASIC](https://github.com/twinbasic/twinbasic).

Initially, this drop-in replacement for `mscorlib.ArrayList` was just to get rid of the `423 MB VMem` overheat added to my VBA projects when using it's `ArrayList` implementation. However, it turned out heavily outperform `mscorlib.ArrayList` practically in every way while also, getting rid of the non-existing memory deallocation after releasing an `ArrayList` instance.


## **Features**

* Less than `0.35 MB VMem` overheat _(vs `423 MB` of `mscorlib`)_
* __Drop-in replacement__ - It's expected to provide the exact same functionality as `mscorlib.ArrayList` when used in `VBA`.
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


## **Perf Tests**



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


