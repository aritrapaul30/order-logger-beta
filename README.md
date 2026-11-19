# Digital Den Order Logger (beta)

Test version of the order logger with two dealers and client payments. The main app stays at `order-logger` until this is approved.

- **Dealer picker:** Rahul Yadav / Sourav Mondal. Each dealer's sheet keeps its own formulas.
- **Orders:** save an order when you place it (status Ordered). When the out-for-delivery message comes, tap the saved row and paste the message to complete it.
- **Out for delivery:** one tap marks Delivered and adds Amount + Commission cut to the client whose card paid.
- **Clients:** what you still owe each client, Pay (cash or "settled via dealer cut"), Share statement image, Mark settled (moves rows to a Settled history tab).
- **Dealer payments** and **Profit** per dealer, with all-dealer totals.

`apps-script/Code.gs` connects the page to the three Google Sheets. Its key and sheet IDs are filled in only inside your own Apps Script project, never in this repository.
