import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { Workbook, SpreadsheetFile } from '@oai/artifact-tool';
process.on('uncaughtException',e=>{console.error(e.stack);process.exit(1);});

const root=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const a=JSON.parse(await fs.readFile(path.join(root,'data/processed/analysis.json'),'utf8'));
const wb=Workbook.create();
const names=['Start_Here','Sales_Overview','Product_Analysis','Customer_RFM','Cohort_Retention','Regional_Analysis','Insights','Data_Quality','Data_Dictionary','Monthly_Data'];
const sh=Object.fromEntries(names.map(n=>[n,wb.worksheets.add(n)]));
const C={navy:'#20364D',teal:'#177D87',blue:'#3157A1',amber:'#FFF0CE',paper:'#F6F8FA',gray:'#627183',line:'#D7DFE7',white:'#FFFFFF'};
const GBP='"£"#,##0.00;("£"#,##0.00);"£"0.00';
const GBPk='"£"0.0,,"m"';
const pct='0.0%';
const col=n=>{let s='';for(n++;n;n=Math.floor((n-1)/26))s=String.fromCharCode(65+(n-1)%26)+s;return s;};
const date=s=>new Date(s+'T00:00:00Z');
const val=(s,cell,v)=>s.getRange(cell).values=[[v]];
const formula=(s,cell,v)=>s.getRange(cell).formulas=[[v]];
function span(s,range,text,opts={}){s.mergeCells(range);const r=s.getRange(range);r.values=[[text]];r.format={font:{name:'Arial',size:opts.size||11,color:opts.color||C.navy,bold:opts.bold||false},wrapText:true,verticalAlignment:'center',...opts};}
function base(s,title,context,end='L',rows=45){
 s.showGridLines=false;s.tabColor=names.indexOf(s.name)<7?C.navy:C.gray;
 s.getRange(`A1:${end}${rows}`).format={font:{name:'Arial',size:11,color:C.navy},rowHeight:23,columnWidth:13,verticalAlignment:'center'};
 s.getRange('A1:A'+rows).format.columnWidth=3;
 span(s,`B2:${end}3`,title,{size:17,bold:true});
 s.getRange(`B3:${end}3`).format.borders={bottom:{style:'thin',color:C.teal}};
 span(s,`B4:${end}4`,context,{size:10,color:C.gray});
}
function table(s,row,headers,rows,name,start=1){
 const end=col(start+headers.length-1),first=col(start),last=row+rows.length;
 s.getRange(`${first}${row}:${end}${last}`).values=[headers,...rows];
 const t=s.tables.add(`${first}${row}:${end}${last}`,true,name);t.style='TableStyleMedium2';
 s.getRange(`${first}${row}:${end}${row}`).format={fill:C.navy,font:{bold:true,color:C.white,name:'Arial',size:11},wrapText:true,rowHeight:34};
 s.getRange(`${first}${row+1}:${end}${last}`).format.rowHeight=22;
 return last;
}
function chart(s,type,ranges,title,from,to,format=GBP){
 const ch=s.charts.add(type,ranges.length===1?s.getRange(ranges[0]):ranges.map(r=>s.getRange(r)));ch.title=title;ch.setPosition(from,to);
 ch.titleTextStyle.fontSize=14;ch.titleTextStyle.typeface='Arial';ch.hasLegend=ranges.length>2;
 ch.legend={position:'bottom',textStyle:{typeface:'Arial',fontSize:11}};
 ch.xAxis={axisType:'textAxis',textStyle:{typeface:'Arial',fontSize:10}};
 ch.yAxis={numberFormatCode:format,numberFormatSourceLinked:false,textStyle:{typeface:'Arial',fontSize:10}};
 ch.series.items.forEach((x,i)=>{x.fill=[C.teal,C.blue,'#B66C3D'][i%3];if(type==='line')x.line={fill:[C.teal,C.blue,'#B66C3D'][i%3],style:'solid',width:2};});
 return ch;
}
function metric(s,range,label,cell,f,fmt=GBP){span(s,range,label,{size:10,color:C.gray});const end=range.split(':')[1].replace(/[0-9]/g,'');const r=Number(cell.match(/[0-9]+/)[0]);s.mergeCells(`${cell}:${end}${r+1}`);formula(s,cell,f);s.getRange(cell).setNumberFormat(fmt);s.getRange(cell).format.font={name:'Arial',size:18,bold:true,color:C.navy};}

// Authoritative input grains: country-month and independent source totals.
{
const s=sh.Monthly_Data;base(s,'Monthly source aggregates','Prepared from every source row. One row per month and country; GBP.','P',a.scope.length+8);
const rows=a.scope.map(x=>[date(x.Month+'-01'),x.Month,x.Country,x.GrossSales,x.ReturnValue,x.Orders]);
table(s,6,['Month start','Month label','Country','Gross sales','Returns','Orders'],rows,'CountryMonth',0);
s.getRange('D7:E'+(rows.length+6)).setNumberFormat(GBP);s.getRange('A7:A'+(rows.length+6)).setNumberFormat('mmm yyyy');s.getRange('A1:B'+(rows.length+6)).format.columnWidth=16;s.getRange('D1:E'+(rows.length+6)).format.columnWidth=18;s.getRange('C1:C'+(rows.length+6)).format.columnWidth=26;
const totals=[['Gross sales',a.totals.gross],['Returns',a.totals.returns],['Net sales',a.totals.net],['Orders',a.totals.orders],['Customers',a.totals.customers],['Repeat customers',a.totals.repeat_customers],['Raw rows',a.totals.raw_rows],['Jan-Nov 2010 net',a.totals.jan_nov['2010']],['Jan-Nov 2011 net',a.totals.jan_nov['2011']],['Identified gross',a.totals.identified_gross],['Unidentified gross',a.totals.unidentified_gross],['Repeated-line net',a.totals.duplicate_net],['Positive-sales codes',a.totals.positive_products],['Codes to reach 80%',a.totals.products_to_80]];
table(s,6,['Independent source control','Value'],totals,'SourceControls',8);s.getRange('I1:I25').format.columnWidth=30;s.getRange('J1:J25').format.columnWidth=22;s.getRange('J7:J20').setNumberFormat('#,##0.00');
span(s,'I23:P25','Controls calculated directly from transaction rows in prepare_data.py. They are independent reconciliation targets, not sums of the displayed analysis tables.',{size:10});
val(s,'L6','Country choices');s.getRange('L7').write([['All countries'],...a.regional.map(x=>[x.Country]).sort((x,y)=>x[0].localeCompare(y[0]))]);
val(s,'O6','Month choices');s.getRange('O7').write(a.monthly.map(x=>[x.Month]));s.getRange('L1:L50').format.columnWidth=27;s.freezePanes.freezeRows(6);
}
const n=a.scope.length+6, sr=`'Monthly_Data'!`;
{
const s=sh.Sales_Overview;base(s,'Retail sales overview','Historical UK online retailer. December 2009–December 2011; GBP.','L',66);
val(s,'B6','Start month');val(s,'E6','End month');val(s,'H6','Country');
val(s,'B7','2009-12');val(s,'E7','2011-12');span(s,'H7:L7','All countries');
for(const c of ['B7','E7']){s.getRange(c).dataValidation={rule:{type:'list',formula1:"'Monthly_Data'!$O$7:$O$31"}};s.getRange(c).format.fill=C.amber;}
s.getRange('H7').dataValidation={rule:{type:'list',formula1:`'Monthly_Data'!$L$7:$L$${a.regional.length+7}`}};s.getRange('H7:L7').format.fill=C.amber;
metric(s,'B9:D9','Gross sales','B10','=SUM(C38:C62)');metric(s,'F9:H9','Returns / adjustments','F10','=SUM(D38:D62)');metric(s,'J9:L9','Net sales','J10','=B10-F10');
metric(s,'B13:D13','Distinct purchase orders','B14','=SUM(F38:F62)','#,##0');metric(s,'F13:H13','Gross average order value','F14','=IF(B14=0,0,B10/B14)');metric(s,'J13:L13','Return value / gross sales','J14','=IF(B10=0,0,F10/B10)',pct);
span(s,'B16:L17','Amber cells filter this dashboard and its chart only. AOV = positive purchase revenue ÷ distinct purchase invoices. Returns are shown separately; December 2011 ends on 9 December.',{size:10});
formula(s,'B18','=IF(B7>E7,"Invalid range: start month is after end month","Monthly net sales for the selected scope")');s.mergeCells('B18:L18');
const rows=a.monthly.map(x=>[x.Month,null,null,null,null,null,x.Month==='2011-12'?'Partial month':'Complete month']);
table(s,37,['Month','Gross sales','Return value','Net sales','Orders','Gross AOV','Coverage'],rows,'MonthlySales');
for(let i=0;i<rows.length;i++){const r=38+i;const filter=`${sr}$B$7:$B$${n},B${r},${sr}$C$7:$C$${n},IF($H$7="All countries","*",$H$7)`;
 for(const [target,src] of [['C','D'],['D','E'],['F','F']])formula(s,target+r,`=IF(AND(B${r}>=$B$7,B${r}<=$E$7),SUMIFS(${sr}$${src}$7:$${src}$${n},${filter}),0)`);
 formula(s,'E'+r,`=C${r}-D${r}`);formula(s,'G'+r,`=IF(F${r}=0,0,C${r}/F${r})`);
}
s.getRange('C38:E62').setNumberFormat(GBP);s.getRange('G38:G62').setNumberFormat(GBP);s.getRange('F38:F62').setNumberFormat('#,##0');s.getRange('C1:G66').format.columnWidth=16;s.getRange('H37:H62').format.columnWidth=22;
chart(s,'line',['B37:B62','E37:E62'],'Monthly net sales (GBP)','B20','L35',GBPk);
span(s,'B64:L66','Months outside the selected date range display zero and are excluded from totals. Order counts are additive here because each invoice belongs to one month and one country, checked during preparation.',{size:10});
}
{
const s=sh.Product_Analysis;base(s,'Product contribution','Gross sales ranking by stock code. Includes service and adjustment codes; no product categories are supplied.','L',a.products.length+37);
metric(s,'B6:E6','Codes producing 80% of gross sales','B7',"='Monthly_Data'!J20",'#,##0');metric(s,'G6:L6','Share of positive-sales codes','G7',"='Monthly_Data'!J20/'Monthly_Data'!J19",pct);
const top=a.products.slice(0,10);table(s,11,['Stock code','Gross sales'],top.map(x=>[x.StockCode,x.GrossSales]),'TopProducts');s.getRange('C12:C21').setNumberFormat(GBP);
chart(s,'bar',['B11:C21'],'Top 10 stock codes by gross sales','F10','L26','"£"0,"k"');
span(s,'B28:L30','Pareto ranks all codes by gross sales. Return-only codes remain in the full table. Descriptions use the last non-missing description for each code. Ranking, shares and cumulative percentages refresh when the preparation script is rerun.',{size:10});
table(s,33,['Rank','Code','Description','Gross sales','Returns','Net sales','Gross share','Cumulative share','Orders'],a.products.map(x=>[x.Rank,x.StockCode,x.Description,x.GrossSales,x.ReturnValue,x.NetSales,x.RevenueShare,x.CumulativeShare,x.Orders]),'ProductDetail');
s.getRange('D1:D'+(a.products.length+34)).format.columnWidth=48;s.getRange('E1:G'+(a.products.length+34)).format.columnWidth=17;s.getRange('E34:G'+(a.products.length+33)).setNumberFormat(GBP);s.getRange('H34:I'+(a.products.length+33)).setNumberFormat(pct);s.getRange('H1:I'+(a.products.length+34)).format.columnWidth=17;s.freezePanes.freezeRows(5);
}
{
const s=sh.Customer_RFM;base(s,'Customer RFM segments','Identified positive purchasers. Fixed reference date: '+a.reference_date+'. Monetary = gross purchases, GBP.','L',a.rfm.length+39);
const end=38+a.rfm.length;
table(s,7,['Segment','Customers','Gross sales','Customer share'],a.segments.map(x=>[x.Segment,null,null,null]),'SegmentSummary');
for(let i=0;i<a.segments.length;i++){let r=8+i;formula(s,'C'+r,`=COUNTIF($L$39:$L$${end},B${r})`);formula(s,'D'+r,`=SUMIF($L$39:$L$${end},B${r},$H$39:$H$${end})`);formula(s,'E'+r,`=C${r}/COUNTA($B$39:$B$${end})`);}
s.getRange('B1:B'+(end+1)).format.columnWidth=25;s.getRange('D8:D13').setNumberFormat(GBP);s.getRange('E8:E13').setNumberFormat(pct);s.getRange('D1:E14').format.columnWidth=19;
chart(s,'bar',['B7:C13'],'Customers by RFM segment','G6','L20','#,##0');
val(s,'B16','Choose segment');val(s,'B17','Champions');s.getRange('B17').format.fill=C.amber;s.getRange('B17').dataValidation={rule:{type:'list',values:a.segments.map(x=>x.Segment)}};
metric(s,'B19:D19','Selected segment customers','B20','=SUMIF(B8:B13,B17,C8:C13)','#,##0');metric(s,'B22:D22','Selected segment gross sales','B23','=SUMIF(B8:B13,B17,D8:D13)');
metric(s,'G22:I22','Repeat purchaser rate','G23',"='Monthly_Data'!J12/'Monthly_Data'!J11",pct);
span(s,'B25:L27','Repeat rate: customers with at least two distinct positive purchase invoices ÷ all identified positive purchasers. Scores use 20th/40th/60th/80th percentile cut points; ties stay together, so segment sizes need not be equal.',{size:10});
span(s,'B28:L30',`Recency cut points: ${a.rfm_cutpoints.RecencyDays.join(', ')} days. Frequency: ${a.rfm_cutpoints.Frequency.join(', ')} orders. Monetary: ${a.rfm_cutpoints.Monetary.map(x=>'£'+x.toFixed(3)).join(', ')}. Recency bands reverse: lower days score higher.`,{size:10});
span(s,'B31:L35','Segment priority: Champions R≥4,F≥4,M≥4; At risk R≤2,F≥3; Loyal R≥3,F≥3; Recent light buyers R≥4,F≤2; Hibernating R≤2; otherwise Developing. The first matching rule wins. Use the detail table filters for country and segment. Country is the latest observed purchase country.',{size:10});
table(s,38,['Customer ID','Country','First observed','Last purchase','Recency days','Orders','Gross spend','R','F','M','Segment'],a.rfm.map(x=>[x.CustomerID,x.Country,new Date(x.FirstPurchase),new Date(x.LastPurchase),x.RecencyDays,x.Frequency,x.Monetary,x.R,x.F,x.M,x.Segment]),'CustomerDetail');
s.getRange('C38:C'+end).format.columnWidth=25;s.getRange('D39:E'+end).setNumberFormat('dd mmm yyyy');s.getRange('D38:E'+end).format.columnWidth=18;s.getRange('H39:H'+end).setNumberFormat(GBP);s.getRange('H38:H'+end).format.columnWidth=18;s.getRange('I38:K'+end).format.columnWidth=7;s.getRange('L1:L'+end).format.columnWidth=25;s.freezePanes.freezeRows(5);
}
{
const s=sh.Cohort_Retention;base(s,'Customer cohort retention','Cohorts use the first observed positive purchase month, not verified customer acquisition.','AB',70);
s.getRange('C1:AB70').format.columnWidth=9;s.getRange('C1:C70').format.columnWidth=13;s.getRange('B1:B70').format.columnWidth=17;
span(s,'B6:AB8','Each rate is distinct cohort customers purchasing at that age ÷ the original cohort size. Blank = not yet observable. Month 0 is 100%. The final observed calendar month, December 2011, is incomplete and is highlighted amber.',{size:11});
table(s,10,['Cohort','Customers',...Array.from({length:25},(_,i)=>'M'+i)],a.cohorts.map(x=>[x.Cohort,x.Customers,...Array(25).fill(null)]),'RetentionRates');
table(s,41,['Cohort','Customers',...Array.from({length:25},(_,i)=>'M'+i)],a.cohorts.map(x=>[x.Cohort,x.Customers,...x.Counts]),'RetentionCounts');
for(let i=0;i<a.cohorts.length;i++){const row=11+i;for(let j=0;j<25;j++){if(a.cohorts[i].Counts[j]!==null)formula(s,col(j+3)+row,`=${col(j+3)}${42+i}/$C${row}`);}const last=a.cohorts[i].Counts.filter(v=>v!==null).length-1;s.getRange(col(last+3)+row).format.borders={bottom:{style:'medium',color:'#DAA746'}};}
s.getRange('D11:AB35').setNumberFormat('0%');s.getRange('D11:AB35').conditionalFormats.add('colorScale',{colors:['#F2F6F8','#98CDD0','#197C85'],thresholds:[0,.5,1]});s.freezePanes.freezeRows(10);s.freezePanes.freezeColumns(3);
}
{
const s=sh.Regional_Analysis;base(s,'Country performance','Customer residence countries. Full observed period; GBP.','L',a.regional.length+39);
table(s,7,['Country','Net sales'],a.regional.slice(1,9).map(x=>[x.Country,x.NetSales]),'InternationalTop');s.getRange('B1:B'+(a.regional.length+39)).format.columnWidth=28;s.getRange('C1:C'+(a.regional.length+39)).format.columnWidth=18;s.getRange('C8:C15').setNumberFormat(GBP);
chart(s,'bar',['B7:C15'],'Largest non-UK markets by net sales','F6','L23','"£"0,"k"');
metric(s,'B18:D18','UK share of net sales','B19',`=SUMIF(B33:B${32+a.regional.length},"United Kingdom",E33:E${32+a.regional.length})/SUM(E33:E${32+a.regional.length})`,pct);
span(s,'B25:L28','The UK is excluded from the chart so smaller markets remain legible; it is included in the full table and every total. Country-level customer counts are not additive because a customer can buy from multiple countries. Orders are positive purchase invoices.',{size:10});
table(s,32,['Country','Gross sales','Returns','Net sales','Orders','Customers','Net share'],a.regional.map(x=>[x.Country,x.GrossSales,x.ReturnValue,x.NetSales,x.Orders,x.IdentifiedCustomers,x.NetSales/a.totals.net]),'CountryDetail');s.getRange('C33:E'+(32+a.regional.length)).setNumberFormat(GBP);s.getRange('C32:E'+(32+a.regional.length)).format.columnWidth=19;s.getRange('H33:H'+(32+a.regional.length)).setNumberFormat(pct);
}
{
const s=sh.Data_Quality;base(s,'Data quality and reconciliation','Cross-sheet import overlap removed. Remaining repeated lines retained. Diagnostic counts may overlap.','E',62);
table(s,6,['Condition','Rows','Treatment'],a.quality,'QualityProfile');s.getRange('B1:B62').format.columnWidth=38;s.getRange('C1:C62').format.columnWidth=17;s.getRange('D1:D62').format.columnWidth=90;s.getRange('D7:D22').format.wrapText=true;s.getRange('B7:D22').format.rowHeight=34;s.getRange('C7:C22').setNumberFormat('#,##0');
const checks=[['Gross sales',`=SUM('Monthly_Data'!D7:D${n})`,"='Monthly_Data'!J7"],['Returns',`=SUM('Monthly_Data'!E7:E${n})`,"='Monthly_Data'!J8"],['Product net',`=SUM('Product_Analysis'!G34:G${33+a.products.length})`,"='Monthly_Data'!J9"],['Country net',`=SUM('Regional_Analysis'!E33:E${32+a.regional.length})`,"='Monthly_Data'!J9"],['Orders',`=SUM('Monthly_Data'!F7:F${n})`,"='Monthly_Data'!J10"],['RFM customers',`=COUNTA('Customer_RFM'!B39:B${38+a.rfm.length})`,"='Monthly_Data'!J11"],['Cohort customers',"=SUM('Cohort_Retention'!C11:C35)","='Monthly_Data'!J11"],['Classified source rows','=C19+C20+C21+C22',"='Monthly_Data'!J13"]];
table(s,25,['Check','Actual','Expected','Difference'],checks.map(x=>[x[0],null,null,null]),'Reconciliations');checks.forEach((x,i)=>{const r=26+i;formula(s,'C'+r,x[1]);formula(s,'D'+r,x[2]);formula(s,'E'+r,`=ROUND(C${r}-D${r},6)`);});s.getRange('C26:E33').setNumberFormat('#,##0.00');s.getRange('E26:E33').conditionalFormats.add('cellIs',{operator:'notEqual',formula:0,format:{fill:'#FDE4DE',font:{color:'#9C3226'}}});
span(s,'B36:E38',`Removing the ${a.quality[3][1].toLocaleString()} repeated rows after their first occurrence would reduce net sales by £${a.totals.duplicate_net.toLocaleString('en-GB',{maximumFractionDigits:2})}. This is a sensitivity case, not an approved cleaning rule.`,{size:11});
span(s,'B40:E43','RFM scores are checked to stay within 1–5. Cohort customer counts cannot exceed their cohort size. Preparation also verifies that purchase invoices map to one month and one country. Automated results and export checks are recorded in docs/validation.json.',{size:11});
span(s,'B45:E48',`${a.totals.overlap_removed.toLocaleString()} records repeated across the overlapping source sheets were removed from the second import. Matches use all eight source fields plus within-sheet occurrence number. ${a.totals.analysis_rows.toLocaleString()} lines remain before monetary exclusions. The removed records are supplied as an audit CSV.`,{size:11});
}
{
const s=sh.Insights;base(s,'Findings and business implications','Descriptive findings from the historical dataset; recommendations are hypotheses to test.','L',49);
const insights=[
 ['Sales and returns',`Gross sales were £${(a.totals.gross/1e6).toFixed(2)}m. Returns/adjustments were ${(a.totals.return_rate*100).toFixed(2)}% of gross, leaving £${(a.totals.net/1e6).toFixed(2)}m net.`, 'Review high-value cancellation invoices before setting a return-reduction target. No refund reason or cost data is available.'],
 ['Repeat purchasing',`${a.totals.repeat_customers.toLocaleString()} of ${a.totals.customers.toLocaleString()} identified purchasers (${(a.totals.repeat_rate*100).toFixed(1)}%) placed at least two positive purchase invoices.`, 'Test retention campaigns with a holdout group. Anonymous purchases are excluded from this customer measure.'],
 ['Revenue concentration',`${a.totals.products_to_80.toLocaleString()} of ${a.totals.positive_products.toLocaleString()} positive-sales stock codes (${(a.totals.products_80_share*100).toFixed(1)}%) produced at least 80% of gross sales.`, 'Prioritize availability checks for leading codes after separating merchandise from service charges in a future enriched catalog.'],
 ['Geographic concentration',`The United Kingdom contributed ${(a.totals.uk_net_share*100).toFixed(1)}% of net sales.`, 'Evaluate the largest international markets separately. Sales alone do not establish profitability or acquisition efficiency.'],
 ['Comparable growth',`Net sales rose ${(a.totals.yoy*100).toFixed(2)}% from Jan–Nov 2010 to Jan–Nov 2011.`, 'Compare matched complete months. December 2011 is partial and should not be compared with full December 2010.']
 ];
for(let i=0;i<insights.length;i++){let r=6+i*8;span(s,`B${r}:L${r}`,`${i+1}. ${insights[i][0]}`,{bold:true,size:12});span(s,`B${r+1}:L${r+3}`,insights[i][1]);span(s,`B${r+4}:L${r+6}`,insights[i][2],{size:10,color:C.gray});}
}
{
const s=sh.Data_Dictionary;base(s,'Data dictionary and methodology','Source: Chen, D. (2012), Online Retail II. UCI Machine Learning Repository.','D',62);
span(s,'B6:D8','Dataset: https://archive.ics.uci.edu/dataset/502/online+retail+ii\nDOI: https://doi.org/10.24432/C5CG6D\nLicense: CC BY 4.0. Downloaded 25 September 2026. Historical coverage: 1 December 2009–9 December 2011.',{size:10});
const fields=[['Invoice → InvoiceNo','Text identifier','Invoice reference; leading C indicates cancellation.'],['StockCode','Text identifier','Product or service/adjustment code. No category or cost information.'],['Description','Text','Item description. Display uses last available description per stock code.'],['Quantity','Numeric units','Signed line quantity. Negative values are return/adjustment candidates.'],['InvoiceDate','Date/time','Timestamp of the invoice line, as supplied; no timezone conversion.'],['Price → UnitPrice','GBP per unit','Source price in pounds sterling. Positive prices required for monetary inclusion.'],['Customer ID → CustomerID','Text identifier','Missing values excluded only from customer analyses.'],['Country','Text','Customer country of residence as supplied.'],['Gross sales','GBP','Quantity × UnitPrice for positive, non-cancelled paid purchase lines.'],['Return value','GBP, positive','Absolute Quantity × UnitPrice for negative quantity or C-prefixed invoices with positive price.'],['Net sales','GBP','Gross sales less return/adjustment value. Not profit.'],['Orders','Distinct invoices','Positive purchase invoices; excludes return-only/cancellation invoices.'],['Gross AOV','GBP per order','Gross sales divided by distinct positive purchase invoices.'],['Recency / frequency / monetary','Days / orders / GBP','Days since latest positive purchase; distinct positive invoices; gross positive purchase spend.'],['Repeat purchase rate','Proportion','Identified customers with ≥2 positive invoices ÷ identified positive purchasers.'],['Cohort retention','Proportion','Distinct active cohort customers at month age ÷ original observed cohort customers.']];
table(s,11,['Field or metric','Unit / type','Definition'],fields,'Dictionary');s.getRange('B1:B62').format.columnWidth=35;s.getRange('C1:C62').format.columnWidth=25;s.getRange('D1:D62').format.columnWidth=90;s.getRange('D12:D27').format.wrapText=true;s.getRange('B12:D27').format.rowHeight=38;
span(s,'B31:D34','Grains: source = observed invoice line; Monthly_Data = month × country; customer table = one identified purchaser; product table = one stock code; regional table = one country. No many-to-many transaction joins are used.',{size:11});
span(s,'B36:D40','The complete source has 1,067,371 rows, exceeding one Excel worksheet. All rows are processed externally, and only clearly labeled aggregates and customer/product detail are embedded. No truncation. No Power Query, PivotTable, slicer, or Data Model is embedded in this version.',{size:11});
span(s,'B42:D45','Amounts retain source precision, including prices with three decimal places. Displayed GBP values round to two decimals. Forecasting is omitted: limited history, a partial final month, and no evaluated holdout model make a forecast unsuitable for this delivery.',{size:11});
}
{
const s=sh.Start_Here;base(s,'Retail Customer Intelligence','Waqas Khan · Data science portfolio project','L',47);
span(s,'B6:L9','Explore sales performance, customer purchasing behavior, product concentration and retention using 1,067,371 genuine historical retail records.',{size:15,bold:true});
const nav=[['Sales_Overview','Change amber month and country cells; review the linked KPIs and chart.'],['Product_Analysis','Inspect gross-sales rankings and Pareto concentration.'],['Customer_RFM','Choose a segment; filter customer detail by country or segment.'],['Cohort_Retention','Read the retention heatmap by first observed purchase month.'],['Regional_Analysis','Compare the UK with international markets.'],['Insights','Use five evidence-based findings for your presentation.'],['Data_Quality','Inspect exclusions, duplicate sensitivity and reconciliation.'],['Data_Dictionary','Read definitions, dataset attribution and limitations.']];
nav.forEach((x,i)=>{const r=12+i*2;val(s,'B'+r,x[0]);s.getRange('B'+r).format.font={name:'Arial',size:11,color:C.blue,bold:true};span(s,`E${r}:L${r+1}`,x[1],{size:11});});s.getRange('B1:B47').format.columnWidth=27;
span(s,'B10:L10','Select the named worksheet tab at the bottom of Excel. Ctrl+Page Up / Page Down moves between sheets.',{size:10,color:C.gray});
span(s,'B30:L33','Refresh: run the preparation script with the official source workbook, then rebuild the Excel workbook. Table edits and dashboard selections update dependent formulas; source transactions are refreshed through the scripts, not Excel Refresh All.',{size:11});
span(s,'B35:L38','Compatibility: created as a macro-free .xlsx with native tables, charts, formulas and dropdowns. Office 2024 64-bit was detected on this laptop. Desktop Excel opening and filter interaction still require a native application check.',{size:11});
span(s,'B40:L44','Start with the sales dashboard, then customer segments and cohort retention. Use the README walkthrough and interview explanation. Data and derived tables are CC BY 4.0; project code is MIT. Source data is historical, not current business performance.',{size:11});
}

// Test controls by changing inputs, observing dependent outputs, then restoring them.
const tests=[];const sales=sh.Sales_Overview;
function near(label,actual,expected,tol=.011){if(typeof actual!=='number'||Math.abs(actual-expected)>tol)throw new Error(`${label}: ${actual} != ${expected}`);tests.push({check:label,actual,expected,passed:true});}
near('Dashboard gross, full scope',sales.getRange('B10').values[0][0],a.totals.gross);
near('Dashboard net, full scope',sales.getRange('J10').values[0][0],a.totals.net);
near('Dashboard orders, full scope',sales.getRange('B14').values[0][0],a.totals.orders,0);
val(sales,'H7','United Kingdom');near('Country control changes gross',sales.getRange('B10').values[0][0],a.regional.find(x=>x.Country==='United Kingdom').GrossSales);
val(sales,'B7','2011-01');val(sales,'E7','2011-11');near('Country and date controls combine',sales.getRange('J10').values[0][0],a.scope.filter(x=>x.Country==='United Kingdom'&&x.Month>='2011-01'&&x.Month<='2011-11').reduce((z,x)=>z+x.NetSales,0));
val(sales,'B7','2011-12');val(sales,'E7','2009-12');near('Reversed date range clears totals',sales.getRange('B10').values[0][0],0,0);
val(sales,'B7','2009-12');val(sales,'E7','2011-12');val(sales,'H7','All countries');
val(sh.Customer_RFM,'B17','At risk');near('Segment control',sh.Customer_RFM.getRange('B20').values[0][0],a.segments.find(x=>x.Segment==='At risk').Customers,0);val(sh.Customer_RFM,'B17','Champions');
wb.recalculate();
const errors=await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!',options:{useRegex:true,maxResults:50},summary:'Formula error scan',maxChars:4000});
console.log('ERROR SCAN',errors.ndjson);
console.log((await wb.inspect({kind:'table',range:'Data_Quality!B25:E33',include:'values,formulas',tableMaxRows:9,tableMaxCols:4,maxChars:4000})).ndjson);
const dest=process.env.WORKBOOK_OUTPUT || path.join(root,'Retail_Customer_Intelligence.xlsx');
await fs.mkdir(path.dirname(dest),{recursive:true});
await (await SpreadsheetFile.exportXlsx(wb)).save(dest);
console.log('EXPORTED',dest);
await fs.mkdir(path.join(root,'assets'),{recursive:true});
const views={Start_Here:'A1:L45',Sales_Overview:'A1:L35',Product_Analysis:'A1:L30',Customer_RFM:'A1:L35',Cohort_Retention:'A1:AB35',Regional_Analysis:'A1:L28',Insights:'A1:L46',Data_Quality:'A1:E48',Data_Dictionary:'A1:D45',Monthly_Data:'A1:P25'};
for(const [name,range] of Object.entries(views)){
try{const blob=await wb.render({sheetName:name,range,scale:1,format:'png'});await fs.writeFile(path.join(root,'assets',name+'.png'),new Uint8Array(await blob.arrayBuffer()));console.log('RENDERED',name);}catch(e){console.error('RENDER FAILED',name,e.message);tests.push({check:'Render '+name,passed:false,error:e.message});}
}
await fs.writeFile(path.join(root,'docs','validation.json'),JSON.stringify({preparation:a.validation,workbook:tests,formula_scan:errors.ndjson,native_excel:'Not yet performed',features:{tables:true,charts:true,dropdowns:true,pivot_tables:false,power_query:false,slicers:false,data_model:false}},null,2));
console.log('DONE');
