# Nebius console billing catalogue

The console's billing pages (Billing > Pricing) list every SKU Nebius sells, in every region, with the date each price took effect and the date each SKU was made. That is Nebius's own price list. It is wider than the docs rate card: it has regions the docs do not list yet, and SKUs that exist at a price of zero before they are sold.

The snapshot is `data/out/console_catalog_2026-10-06.csv`: 2,401 SKUs in 20 regions, all current prices, none scheduled. The region table is `data/out/console_regions.csv`.

## It needs a login

The catalogue sits behind a console login. There is no public endpoint. This repo does not log in, and it stores no cookie, token, or password. The pages are saved from a browser that is already logged in, and `scripts/build_console_catalog.py` reads the saved file.

The billing gateway is `POST https://console.nebius.com/api-mfe/billing/gateway/root/<service>/<method>`, with a JSON body. Each call needs two headers: `x-csrf-token`, which the console page holds in `window.__DATA__.csrfToken`, and `X-Requested-With: XMLHttpRequest`. Without the token the gateway answers 403 `EBADCSRFTOKEN`. The token belongs to the session. It is read when the snippet runs and is never written down.

To save the catalogue, open `console.nebius.com`, log in, open the developer console, and paste:

```js
const call = async (path, body) => {
  const response = await fetch('/api-mfe/billing/gateway/root/' + path, {
    method: 'POST',
    headers: {
      'content-type': 'application/json',
      'x-csrf-token': window.__DATA__.csrfToken,
      'X-Requested-With': 'XMLHttpRequest',
    },
    body: JSON.stringify(body),
  });
  return response.json();
};
let token = '';
const pages = [];
do {
  const page = await call('customers/listFullView', {filters: {}, additionalData: {}, pageSize: 500, pageToken: token});
  pages.push(page);
  token = page.nextPageToken || '';
} while (token && pages.length < 40);
const link = document.createElement('a');
link.href = URL.createObjectURL(new Blob([JSON.stringify(pages)], {type: 'application/json'}));
link.download = 'nebius_list_full_view.json';
link.click();
```

The browser saves `nebius_list_full_view.json`. Five pages hold the whole catalogue. Then:

```
python scripts/build_console_catalog.py nebius_list_full_view.json --date 2026-10-06
```

That writes `data/out/console_catalog_<date>.csv` and rebuilds `console_regions.csv`. The saved JSON has no account data in it, only SKUs and prices, but keep it out of the repo anyway. `data/cache/` is ignored.

Three other gateway calls answer the same way. `pricingPolicy/listSkuEffectivePriceSeries` is the spot price history in 15-minute, 1-hour, or 2-hour steps, and it is empty until spot starts on 8 Oct 2026. `customers/listAvailablePolicyPlatforms` takes a project id and lists the spot SKUs with their bid bounds (B300 $0.99 to $9.49). `pricingPolicy/getSkusSpotPrices` takes SKU ids. They are not scraped here.

## Columns

`effective_from` is when the price took effect. `sku_created` is when the SKU record was made. `price_created` is when the price record was made, which can be days before it takes effect: the 1 Oct 2026 H100 price of $4.50 was made on 29 Sep. `price_usd` is per `pricing_unit` (a GPU, a core, a GiB). A price of zero means the SKU exists and is not sold yet.

## Regions

`console_regions.csv` has one row per region, in the order the first SKU appeared.

| Region | First SKU | Place | Basis |
|---|---|---|---|
| eu-north1 | 9 Sep 2024 | Finland | docs |
| eu-west1 | 13 Nov 2024 | France | docs |
| eu-north2 | 17 Feb 2025 | Iceland | docs |
| us-central1 | 18 Mar 2025 | Kansas City | docs |
| us-east1 | 16 Sep 2025 | Vineland, inferred | not on the docs page |
| me-west1 | 3 Oct 2025 | Israel | docs |
| uk-south1 | 14 Oct 2025 | United Kingdom | docs |
| eu-west2 | 27 Mar 2026 | France | docs |
| us-south1, eu-west3 | 28 May 2026 | Oklahoma City, Béthune, inferred | not on the docs page |
| uk-south2, eu-south1, us-north1 | 28 May 2026 | UK, Madrid, Woodbury | docs |
| me-west2 | 12 Jun 2026 | Masmiyya, inferred | not on the docs page |
| me-west3, eu-north3, in-south1 | 16 Sep 2026 | Beit Shemesh, Mäntsälä, India | not on the docs page |
| us-central2, us-south2 | 22 Sep 2026 | Independence, Birmingham, inferred | not on the docs page |
| eu-north4 | 24 Sep 2026 | Hüüru, inferred | not on the docs page |

The ten regions on `docs.nebius.com/overview/regions` are all in the catalogue, and their place comes from `data/region_crosswalk.json`, unchanged. The other ten are not on the docs page. Their place is an inference, kept in a separate file, `data/console_region_inferences.json`, with the reason for each one.

A first SKU date is when the region's first SKU record was made. It is not a launch date. Five regions share 28 May 2026, and on that day each got only GPU SKUs, three days before the 1 Jun 2026 price change. One date across five regions looks like a batch of records, not five launches. Read it as "by 28 May".

Vera Rubin is in six regions. eu-north1 got Vera Rubin VM SKUs on 20 Aug 2026. eu-north3, in-south1, and me-west3 followed on 16 Sep, then us-central2 and us-south2 on 22 Sep. All of them are priced at zero. us-east1 has only bare-metal GB200 and GB300 servers, and no VM a customer can start, which fits a single-customer build.
