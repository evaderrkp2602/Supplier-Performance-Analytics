# Supplier Performance Analytics

Purchasing analytics on real public data: delivery performance,
price variance and quality. Tools: Excel, SQL, Python, Power BI.

## Questions this project answers
1. Which vendors deliver late, and by how many days?
2. Which vendors and items are priced above their own baseline?
3. What does an SPC chart say about a process's defect rate?

## Metric definitions (written before looking at the data)
- On-time delivery (OTD): delivered date <= scheduled date
- Lead time: delivered date minus PO sent date
- Price variance: unit price vs. that item's prior-period average
- Landed cost: unit price + allocated freight, compared within the same product group
- Minimum shipments: at least 10 per vendor before a vendor is ranked

## Data sources
(To be filled in as I download each one, with link and license.)
po_sent_to_vendor_date: 5,732 rows not a valid date: those rows say “N/A - From RDC” or “Date Not Captured” instead of a date, so the script left them empty. You can’t calculate lead time for them.
freight_cost: 4,126 rows are text: freight is a note like “Freight Included in Commodity Cost” instead of a dollar amount. The script keeps the text in its own column.
freight_included: 1,442 rows: how many of those notes say freight is already inside the price. The script treats those as $0 freight for landed cost, which is a rule you chose.
Overall on-time delivery: 88.5% is your first real metric: out of 10,324 shipments, about 88.5% arrived on or before the scheduled date.
Saved data/scms.db, table 'shipments' means the cleaned data now lives in a database file. Next you’ll ask it questions in SQL, instead of scrolling a spreadsheet.
