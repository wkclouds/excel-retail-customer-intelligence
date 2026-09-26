"""Profile all rows; remove proven cross-sheet import overlap, retain within-sheet repeats."""
from pathlib import Path
import argparse, hashlib, json, zipfile, urllib.request
from datetime import datetime, timezone
import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
URL = 'https://archive.ics.uci.edu/static/public/502/online%2Bretail%2Bii.zip'

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--download', action='store_true')
    args = parser.parse_args()
    raw = ROOT / 'data/raw'
    out = ROOT / 'data/processed'
    raw.mkdir(parents=True, exist_ok=True)
    out.mkdir(parents=True, exist_ok=True)
    source = raw / 'online_retail_II.xlsx'
    if args.download:
        archive = raw / 'online-retail-ii.zip'
        urllib.request.urlretrieve(URL, archive)
        with zipfile.ZipFile(archive) as z:
            with z.open('online_retail_II.xlsx') as f:
                source.write_bytes(f.read())
    if not source.exists():
        raise SystemExit('Download the official ZIP and extract online_retail_II.xlsx into data/raw, or pass --download.')
    sheets = pd.read_excel(source, sheet_name=None, engine='openpyxl')
    print('Loaded source sheets:', {k: len(v) for k,v in sheets.items()}, flush=True)
    original_columns = list(next(iter(sheets.values())).columns)
    sheet_profile = {name: {'rows':len(frame),'start':str(frame.InvoiceDate.min()),'end':str(frame.InvoiceDate.max()),'within_sheet_repeated_rows':int(frame.duplicated().sum())} for name,frame in sheets.items()}
    raw_frame = pd.concat(sheets.values(), ignore_index=True)
    raw_count = len(raw_frame)
    assert raw_count == 1067371, 'Unexpected source row count: review provenance before proceeding.'
    raw_repeats = int(raw_frame.duplicated().sum())
    frames=list(sheets.values())
    first=frames[0].copy()
    second=frames[1].copy()
    # Match every original field AND occurrence number. This removes only the
    # repeated import across source sheets while preserving repeated lines inside a sheet.
    first['_occurrence']=first.groupby(original_columns,dropna=False).cumcount()
    second['_occurrence']=second.groupby(original_columns,dropna=False).cumcount()
    joined=second.merge(first,on=original_columns+['_occurrence'],how='left',indicator=True,validate='one_to_one')
    cross_sheet=joined['_merge'].eq('both')
    removed=joined.loc[cross_sheet,original_columns]
    overlap_count=int(cross_sheet.sum())
    removed.to_csv(out/'cross_sheet_overlap_removed.csv',index=False)
    df=pd.concat([frames[0],joined.loc[~cross_sheet,original_columns]],ignore_index=True)
    duplicate = df.duplicated(keep='first')
    df = df.rename(columns={'Invoice':'InvoiceNo', 'Price':'UnitPrice', 'Customer ID':'CustomerID'})
    for c in ['InvoiceNo', 'StockCode', 'Country']:
        df[c] = df[c].astype('string').str.strip()
    df['InvoiceDate'] = pd.to_datetime(df['InvoiceDate'], errors='coerce')
    df['CustomerID'] = pd.to_numeric(df['CustomerID'], errors='coerce').astype('Int64').astype('string')
    df['Month'] = df.InvoiceDate.dt.strftime('%Y-%m')
    cancelled = df.InvoiceNo.str.upper().str.startswith('C').fillna(False)
    valid_price = df.UnitPrice.gt(0)
    valid_keys = df.InvoiceDate.notna() & df.InvoiceNo.notna() & df.Country.notna()
    purchase = valid_keys & valid_price & df.Quantity.gt(0) & ~cancelled
    returns = valid_keys & valid_price & (df.Quantity.lt(0) | cancelled) & df.Quantity.ne(0)
    df['GrossSales'] = np.where(purchase, df.Quantity * df.UnitPrice, 0.0)
    df['ReturnValue'] = np.where(returns, abs(df.Quantity * df.UnitPrice), 0.0)
    df['NetSales'] = df.GrossSales - df.ReturnValue
    df['PurchasedUnits'] = np.where(purchase, df.Quantity, 0)
    df['ReturnedUnits'] = np.where(returns, abs(df.Quantity), 0)
    pos = df.loc[purchase].copy()
    # Country and month must be consistent within an invoice for additive order counts.
    invoice_check = pos.groupby('InvoiceNo').agg(months=('Month','nunique'), countries=('Country','nunique'))
    assert invoice_check.months.max() == 1 and invoice_check.countries.max() == 1
    def grouped(keys):
        g = df.groupby(keys, dropna=False)[['GrossSales','ReturnValue','NetSales','PurchasedUnits','ReturnedUnits']].sum()
        g['Orders'] = pos.groupby(keys).InvoiceNo.nunique()
        g['Orders'] = g.Orders.fillna(0).astype(int)
        return g.reset_index()
    monthly = grouped(['Month'])
    scope = grouped(['Month','Country'])
    regional = grouped(['Country']).sort_values('NetSales', ascending=False)
    regional['IdentifiedCustomers'] = regional.Country.map(pos.groupby('Country').CustomerID.nunique()).fillna(0).astype(int)
    products = grouped(['StockCode']).sort_values('GrossSales', ascending=False)
    descriptions = df.dropna(subset=['Description']).groupby('StockCode').Description.last()
    products.insert(1, 'Description', products.StockCode.map(descriptions).fillna('Description unavailable'))
    products['RevenueShare'] = products.GrossSales / products.GrossSales.sum()
    products['CumulativeShare'] = products.RevenueShare.cumsum()
    products['Rank'] = range(1,len(products)+1)
    identified = pos.dropna(subset=['CustomerID'])
    ref = df.InvoiceDate.max().normalize() + pd.Timedelta(days=1)
    rfm = identified.groupby('CustomerID').agg(FirstPurchase=('InvoiceDate','min'),LastPurchase=('InvoiceDate','max'),Frequency=('InvoiceNo','nunique'),Monetary=('GrossSales','sum'))
    rfm['RecencyDays'] = (ref-rfm.LastPurchase.dt.normalize()).dt.days
    # Percentile cut points keep identical observed values in the same band.
    cuts = {col: rfm[col].quantile([.2,.4,.6,.8]).tolist() for col in ['RecencyDays','Frequency','Monetary']}
    for col, score in [('RecencyDays','R'),('Frequency','F'),('Monetary','M')]:
        band = np.searchsorted(cuts[col], rfm[col].to_numpy(), side='left') + 1
        rfm[score] = 6-band if col=='RecencyDays' else band
    rfm['Segment'] = np.select([
        (rfm.R>=4)&(rfm.F>=4)&(rfm.M>=4),
        (rfm.R<=2)&(rfm.F>=3),
        (rfm.R>=3)&(rfm.F>=3),
        (rfm.R>=4)&(rfm.F<=2),
        rfm.R<=2,
    ], ['Champions','At risk','Loyal','Recent light buyers','Hibernating'], default='Developing')
    rfm['Country'] = identified.sort_values('InvoiceDate').groupby('CustomerID').Country.last()
    rfm['Cohort'] = rfm.FirstPurchase.dt.strftime('%Y-%m')
    rfm = rfm.reset_index()
    segments = rfm.groupby('Segment').agg(Customers=('CustomerID','size'),GrossSales=('Monetary','sum'),AverageRecency=('RecencyDays','mean'),AverageFrequency=('Frequency','mean')).reset_index()
    activity = identified[['CustomerID','Month']].drop_duplicates().merge(rfm[['CustomerID','Cohort']],on='CustomerID',validate='many_to_one')
    activity['Age'] = (pd.to_datetime(activity.Month).dt.year-pd.to_datetime(activity.Cohort).dt.year)*12 + pd.to_datetime(activity.Month).dt.month-pd.to_datetime(activity.Cohort).dt.month
    counts = activity.groupby(['Cohort','Age']).CustomerID.nunique().unstack()
    last_month = pd.Period(df.InvoiceDate.max(), freq='M')
    cohort_rows=[]
    for cohort,row in counts.iterrows():
        max_age = (last_month-pd.Period(cohort,freq='M')).n
        vals = [int(row.get(age,0)) if pd.notna(row.get(age,0)) else 0 for age in range(max_age+1)]
        cohort_rows.append({'Cohort':cohort,'Customers':vals[0],'Counts':vals+[None]*(25-len(vals))})
    # Same complete calendar months only; December 2011 is incomplete.
    jan_nov = {str(y): float(monthly.loc[monthly.Month.between(f'{y}-01',f'{y}-11'),'NetSales'].sum()) for y in [2010,2011]}
    quality = [
        ['Original source rows',raw_count,'Every source row profiled; exact cross-sheet import overlap removed separately.'],
        ['Missing customer ID',int(df.CustomerID.isna().sum()),'Included in valid sales; excluded from customer, RFM and cohort calculations.'],
        ['Missing description',int(df.Description.isna().sum()),'Keep product code; use an available description or an explicit unavailable label.'],
        ['Within-population repeated lines',int(duplicate.sum()),'Retained after removing cross-sheet overlap. No line identifier proves these are errors.'],
        ['Cancellation prefix',int(cancelled.sum()),'InvoiceNo starts C, case insensitive.'],
        ['Negative quantity',int(df.Quantity.lt(0).sum()),'Return/adjustment candidate; requires positive unit price for monetary inclusion.'],
        ['Zero unit price',int(df.UnitPrice.eq(0).sum()),'Excluded from monetary metrics; quantity not treated as paid demand.'],
        ['Negative unit price',int(df.UnitPrice.lt(0).sum()),'Excluded from monetary metrics.'],
        ['Zero quantity',int(df.Quantity.eq(0).sum()),'Excluded from sales and return populations.'],
        ['Invalid date/invoice/country',int((~valid_keys).sum()),'Excluded from monetary and order populations.'],
        ['Positive-quantity cancellations',int((cancelled&df.Quantity.gt(0)).sum()),'If present, classified as returns at absolute line value.'],
        ['Negative quantity without C prefix',int((~cancelled&df.Quantity.lt(0)).sum()),'Treated as return/adjustment if price is positive; intent cannot be proven.'],
        ['Included purchase lines',int(purchase.sum()),'Positive quantity and price, valid keys, no cancellation prefix.'],
        ['Included return/adjustment lines',int(returns.sum()),'Positive price, nonzero quantity, valid keys, and negative quantity or C prefix.'],
        ['Excluded monetary lines',int((~(purchase|returns)).sum()),'Reconciles included purchase + return + excluded lines to source.'],
        ['Cross-sheet overlap removed',overlap_count,'Exact eight-field + occurrence matches imported again in the second sheet.'],
    ]
    totals = dict(gross=float(df.GrossSales.sum()),returns=float(df.ReturnValue.sum()),net=float(df.NetSales.sum()),orders=int(pos.InvoiceNo.nunique()),customers=int(identified.CustomerID.nunique()),repeat_customers=int((rfm.Frequency>=2).sum()),raw_rows=raw_count,analysis_rows=len(df),overlap_removed=overlap_count,purchase_lines=int(purchase.sum()),return_lines=int(returns.sum()),excluded_lines=int((~(purchase|returns)).sum()),duplicate_net=float(df.loc[duplicate,'NetSales'].sum()),unidentified_gross=float(df.loc[df.CustomerID.isna(),'GrossSales'].sum()),identified_gross=float(rfm.Monetary.sum()),yoy=jan_nov['2011']/jan_nov['2010']-1,jan_nov=jan_nov)
    totals['aov']=totals['gross']/totals['orders']
    totals['repeat_rate']=totals['repeat_customers']/totals['customers']
    totals['return_rate']=totals['returns']/totals['gross']
    totals['uk_net_share']=float(regional.loc[regional.Country.eq('United Kingdom'),'NetSales'].iloc[0]/totals['net'])
    totals['products_to_80']=int((products.CumulativeShare<.8).sum()+1)
    totals['positive_products']=int(products.GrossSales.gt(0).sum())
    totals['products_80_share']=totals['products_to_80']/totals['positive_products']
    overlap_duplicates = int(duplicate.sum()) - sum(x['within_sheet_repeated_rows'] for x in sheet_profile.values())
    assertions=[]
    for label,value,expected in [('Monthly net',monthly.NetSales.sum(),totals['net']),('Country net',regional.NetSales.sum(),totals['net']),('Product net',products.NetSales.sum(),totals['net']),('Country-month net',scope.NetSales.sum(),totals['net']),('Monthly orders',monthly.Orders.sum(),totals['orders']),('Cohort customers',sum(x['Customers'] for x in cohort_rows),totals['customers'])]:
        assert abs(value-expected)<.01,(label,value,expected)
        assertions.append({'check':label,'actual':float(value),'expected':float(expected),'passed':True})
    assert all(rfm[c].between(1,5).all() for c in ['R','F','M'])
    assert all(0<=v<=row['Customers'] for row in cohort_rows for v in row['Counts'] if v is not None)
    assert totals['purchase_lines']+totals['return_lines']+totals['excluded_lines']+overlap_count==raw_count
    for name,frame in [('monthly',monthly),('country_month',scope),('regional',regional),('products',products),('customers_rfm',rfm),('segments',segments)]:
        frame.to_csv(out/f'{name}.csv',index=False)
    payload = {'totals':totals,'monthly':monthly.to_dict('records'),'scope':scope.to_dict('records'),'regional':regional.to_dict('records'),'products':products.to_dict('records'),'rfm':json.loads(rfm.to_json(orient='records',date_format='iso')),'segments':segments.to_dict('records'),'cohorts':cohort_rows,'quality':quality,'rfm_cutpoints':cuts,'reference_date':str(ref.date()),'start_date':str(df.InvoiceDate.min()),'end_date':str(df.InvoiceDate.max()),'source_sheets':{k:len(v) for k,v in sheets.items()},'original_columns':original_columns,'source_sha256':hashlib.sha256(source.read_bytes()).hexdigest(),'prepared_utc':datetime.now(timezone.utc).isoformat(),'download_date':'2026-09-25','validation':assertions}
    (out/'analysis.json').write_text(json.dumps(payload,indent=2,allow_nan=False),encoding='utf-8')
    (out/'source_profile.json').write_text(json.dumps({'sheets':sheet_profile,'raw_repeated_rows':raw_repeats,'cross_sheet_overlap_removed':overlap_count,'retained_repeated_rows':int(duplicate.sum()),'analysis_rows':len(df),'matching_rule':'All eight source fields plus within-sheet occurrence number'},indent=2),encoding='utf-8')
    print(json.dumps({k:v for k,v in payload.items() if k in ['totals','quality','rfm_cutpoints','source_sha256','validation']},indent=2),flush=True)

if __name__=='__main__':
    main()
