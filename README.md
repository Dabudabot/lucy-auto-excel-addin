# Lucy automatization Excel AddIn
Small excel addin for automatization for statistics and charts

How it looks:

![howitlooks](docs/howitlooks.PNG)

How to use:

In case if you have for example such table:

![example1](docs/example1.PNG)

You fullfill the addin with following data:

![example1](docs/example11.PNG)

Sheets range - on current page addin will look throught this range and find each sheet in this document from this range.

Cells range - on found sheets it will look for text that equals to the data from this range in current page.

Jump amount - text from previous field may go with gaps due to target values (in the example above it is options), this number is amount of them

Append - if checked add to target lists new row else replace the last one

Date suffix - on Y axis replaced or added values will have this suffix

The result of the work:

![example1](docs/example12.PNG)

In yellow you see new data and charts are moved the series.

Another example:

![example1](docs/example2.PNG)

Here we do not have target values after cell

![example1](docs/example21.PNG)

Result

![example1](docs/example22.PNG)