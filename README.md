# arrowExcel

Arrow Excel is a simple VBA library intended to mimic the -> threading macro from Clojure, along with a couple of auxiliary functions (Split, Partition, First, Second, Nth, Filter, Map (sort of, don't get excited)).  I primarily wrote it for the purpose of parsing multi-column datasets downloaded into Excel via the Bloomberg Excel plugin.  

##### Motivation ##### 
Bloomberg's Excel plugin is a little wonky.  One can choose between a) retrieving a single datapoint, which behaves as expected and can be used as the input to another function (within the same cell), and b) retrieving an entire table, whereby the plugin will write a table out over multiple cells.  It is difficult to use the resulting dataset because it can be of variable length, and is limiting, because the formula itself cannot be nested inside another formula. 

To get around this I've begun to use the "Aggregate=YES" flag so that the entire table shows up as a single cell value.  The table below of identifer types and identifiers would normally be laid out in a 2x2 grid:

<pre><code>"REGS"  "XS1234567890"
"BBG"   "34534322"</code></pre>

With the flag on, it's all in one cell:

<pre><code>"REGS XS1234567890 BBG 34534322"</code></pre>

Once it's in this form, I can use my Arrow Excel formulas to extract the value(s) I desire.


##### Usage #####
Assume a cell A1 with contents "REGS XS1234567890 BBG 34534322".

<pre><code>=arrow(A1, aSplit(" "), aPartition(2), aFilter(aEquals(aFirst(), "REGS")), aMap(aSecond()))</code></pre>

will yield "XS1234567890"


##### Install #####
Copy the code and make a new xla file with it.  

##### What's not to like? ##### 
- No comments, no error handling, and no tests, so it's got that going for it.
- Works only with the auxiliary functions I have written specifically with this in mind (Greenspun's Tenth?)
- It's not exactly modular, and has one big Select Case statement to do the dispatching
- It's written in VBA, the world's most painful language

See my other projects for better stuff.
