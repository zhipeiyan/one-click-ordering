# One click ordering
Some shopping websites powered by Ecwid share similar ordering systems. Users can group order with their friends to save the shipping cost.

This script will use a `.xlsx` file as input and put all the items in the cart in one click.

## Usage
Write down all ordres of each user on any Excel-like file. One order per sheet. [Here](https://m9gedv4gqs.larksuite.com/sheets/shtusOAsNn8b5OurYfpnswP0AQb)'s a sample group order sheet. Just export the sheet to .xlsx format.

The order sheet should have every three columns as `Item`, `Quantity` and `Unit price` for each user and the item should be the link to the item page. Type the quantity and unit price manually.

You may have to download the browser [driver](https://www.selenium.dev/documentation/en/webdriver/driver_requirements/#quick-reference) to your computer.

Adjust the configurations like file paths in the script. And run the script.
