# **ArrayList**
`VBA` / `VB6` mscorlib's **ArrayList drop-in replacement**, implemented using [twinBASIC](https://github.com/twinbasic/twinbasic).

Initially, this drop-in replacement for `mscorlib.ArrayList` was just to get rid of the `423 MB VMem` overheat added to my VBA projects when using it's `ArrayList` implementation. However, it turned out heavily outperform `mscorlib.ArrayList` practically in every way while also, getting rid of the non-existing memory deallocation after releasing an `ArrayList` instance.


## **Features**

* Less than `0.35 MB VMem` overheat _(vs `423 MB` of `mscorlib`)_
* __Drop-in replacement__ - Is expected to provide the exact same functionality as `mscorlib.ArrayList` in order to easily replace it.
* Unlike `mscorlib.ArrayList`, seamlessly accepts `VBA Array`s or any other enumerable type as parameter.
* Provides an advanced `Enumerator` allowing the use of `For Each` within subranges, backwards enumeration, custom iteration steps and direct access to the backing enumerator instance allowing an even wider set of possibilities while iterating the `Enumerator`.
* The [`Enumerator`](https://github.com/Theadd/ArrayList/blob/main/ArrayListLib/Sources/Enumerator.twin#L21) class is publicly accessible so you can reuse it anywhere else in your code.


## **Performance Comparison**

<table>
    <thead>
        <tr>
            <th> </th>
            <th colspan=2>ArrayList (DLL)</th>
            <th colspan=2>ArrayList (VBA)</th>
            <th colspan=2>mscorlib.ArrayList</th>
            <th colspan=2>VBA Collection</th>
        </tr>
    </thead>
    <tbody>
        <tr>
            <th>.Add (x5000)</th>
            <td>0s</td>
            <td><code>x1</code></td>
            <td>0s</td>
            <td><code>x1</code></td>
            <td>0.031s</td>
            <td><code>x31</code></td>
            <td>0s</td>
            <td><code>x1</code></td>
        </tr>
        <tr>
            <th>.Add (250x5000)</th>
            <td>0.156s</td>
            <td><code>x1</code></td>
            <td>0.195s</td>
            <td><code>x1.3</code></td>
            <td>3.578s</td>
            <td><code>x22.9</code></td>
            <td>0.156s</td>
            <td><code>x1</code></td>
        </tr>
        <tr>
            <th>.Clone (50x5000)</th>
            <td>0s</td>
            <td><code>x1</code></td>
            <td>0.047s</td>
            <td><code>x47</code></td>
            <td>0s</td>
            <td><code>x1</code></td>
            <td>2.945s</td>
            <td><code>x2945</code></td>
        </tr>
        <tr>
            <th>.Insert Index:=0 (50x1000)</th>
            <td>0.727s</td>
            <td><code>x90.9</code></td>
            <td>0.914s</td>
            <td><code>x114.3</code></td>
            <td>1.008s</td>
            <td><code>x126</code></td>
            <td>0.008s</td>
            <td><code>x1</code></td>
        </tr>
        <tr>
            <th>.Insert Index:=RND (10x5000)</th>
            <td>0.82s</td>
            <td><code>x1</code></td>
            <td>1.008s</td>
            <td><code>x1.2</code></td>
            <td>1.086s</td>
            <td><code>x1.3</code></td>
            <td>0.898s</td>
            <td><code>x1.1</code></td>
        </tr>
        <tr>
            <th>.Item (Read) SEQ (10x5000)</th>
            <td>0s</td>
            <td><code>x1</code></td>
            <td>0.008s</td>
            <td><code>x8</code></td>
            <td>0.227s</td>
            <td><code>x227</code></td>
            <td>0.562s</td>
            <td><code>x562</code></td>
        </tr>
        <tr>
            <th>.Item (Read) SEQ+RND (20x5000)</th>
            <td>0.016s</td>
            <td><code>x1</code></td>
            <td>0.039s</td>
            <td><code>x2.4</code></td>
            <td>0.703s</td>
            <td><code>x43.9</code></td>
            <td>2.367s</td>
            <td><code>x147.9</code></td>
        </tr>
        <tr>
            <th>.RemoveAt Index:=0 (x5000)</th>
            <td>0.008s</td>
            <td><code>x8</code></td>
            <td>0.039s</td>
            <td><code>x39</code></td>
            <td>0.023s</td>
            <td><code>x23</code></td>
            <td>0s</td>
            <td><code>x1</code></td>
        </tr>
        <tr>
            <th>.RemoveAt Index:=RND (x12000)</th>
            <td>0.344s</td>
            <td><code>x1.7</code></td>
            <td>0.477s</td>
            <td><code>x2.3</code></td>
            <td>0.203s</td>
            <td><code>x1</code></td>
            <td>0.914s</td>
            <td><code>x4.5</code></td>
        </tr>
        <tr>
            <th>.RemoveAt Index:=LAST (x12000)</th>
            <td>0s</td>
            <td><code>x1</code></td>
            <td>0.031s</td>
            <td><code>x31</code></td>
            <td>0.031s</td>
            <td><code>x31</code></td>
            <td>0.477s</td>
            <td><code>x477</code></td>
        </tr>
        <tr>
            <th>.AddRange Array(24000)</th>
            <td>0s</td>
            <td><code>x1</code></td>
            <td>0s</td>
            <td><code>x1</code></td>
            <td>0.086s</td>
            <td><code>x86</code></td>
            <td>0s</td>
            <td><code>x1</code></td>
        </tr>
        <tr>
            <th>.AddRange Array(15000)</th>
            <td>0s</td>
            <td><code>x1</code></td>
            <td>0s</td>
            <td><code>x1</code></td>
            <td>0.039s</td>
            <td><code>x39</code></td>
            <td>0s</td>
            <td><code>x1</code></td>
        </tr>
        <tr>
            <th>.GetRange Index:=RND (1000x5000)</th>
            <td>0s</td>
            <td><code>x1</code></td>
            <td>0s</td>
            <td><code>x1</code></td>
            <td>0.031s</td>
            <td><code>x31</code></td>
            <td>N/A</td>
            <td></td>
        </tr>
        <tr>
            <th>.AddRange Range (x1000)</th>
            <td>0.07s</td>
            <td><code>x1</code></td>
            <td>0.078s</td>
            <td><code>x1.1</code></td>
            <td>0.211s</td>
            <td><code>x3</code></td>
            <td>N/A</td>
            <td></td>
        </tr>
        <tr>
            <th>For Each (10x100000)</th>
            <td>0.109s</td>
            <td><code>x4.7</code></td>
            <td>3.664s</td>
            <td><code>x159.3</code></td>
            <td>0.867s</td>
            <td><code>x37.7</code></td>
            <td>0.023s</td>
            <td><code>x1</code></td>
        </tr>
        <tr>
            <th>For Each In .Items (10x100000)</th>
            <td>0.016s</td>
            <td><code>x1</code></td>
            <td>0.017s</td>
            <td><code>x1.1</code></td>
            <td>0.867s</td>
            <td><code>x54.2</code></td>
            <td>0.023s</td>
            <td><code>x1.4</code></td>
        </tr>
        <tr>
            <th>.Sort (100000)</th>
            <td>0.219s</td>
            <td><code>x3.5</code></td>
            <td>0.234s</td>
            <td><code>x3.8</code></td>
            <td>0.062s</td>
            <td><code>x1</code></td>
            <td>N/A</td>
            <td></td>
        </tr>
        <tr>
            <th>.Sort w/ Comparer (100000)</th>
            <td>0.492s</td>
            <td><code>x1</code></td>
            <td>0.781s</td>
            <td><code>x1.6</code></td>
            <td>1.195s</td>
            <td><code>x2.4</code></td>
            <td>N/A</td>
            <td></td>
        </tr>
    </tbody>
</table>

