import os
import numpy as np
import pandas as pd
import xlsxwriter

os.chdir('/home/renku/work')
SRC = 'input-data/Project 1 - DTB/bioreaction_remaining.csv'
OUT = 'DTB_outputs/DTB_batches_charts.xlsx'
os.makedirs(os.path.dirname(OUT), exist_ok=True)

df = pd.read_csv(SRC)
df['to_predict'] = df['to_predict'].astype(str).str.lower().eq('true')
T = ['vcd', 'tcd', 'substrate', 'byproduct']
F = ['feed_substrate', 'feed_b', 'feed_c', 'feed_d']
df['bn'] = df.batch_id.str.extract(r'(\d+)')[0].astype(int)
df = df.sort_values(['bn', 'elapsed_time'])
batches = list(dict.fromkeys(df.batch_id))
test = set(df.loc[df.to_predict, 'batch_id'])

TLAB = {'vcd': 'VCD (viable cell density)', 'tcd': 'TCD (total cell density)',
        'substrate': 'Substrate', 'byproduct': 'Byproduct'}
COL = {'vcd': '1F77B4', 'tcd': '9467BD', 'substrate': '2CA02C', 'byproduct': 'D62728',
       'volume': '7F7F7F', 'feed_substrate': '17BECF', 'feed_b': 'FF7F0E', 'feed_c': '8C564B', 'feed_d': 'E377C2'}

wb = xlsxwriter.Workbook(OUT, {'nan_inf_to_errors': True})
f_title = wb.add_format({'bold': True, 'font_size': 14, 'font_color': '1F4E79'})
f_sub = wb.add_format({'italic': True, 'font_color': '595959'})
f_hdr = wb.add_format({'bold': True, 'bg_color': '1F4E79', 'font_color': 'FFFFFF', 'border': 1, 'text_wrap': True, 'valign': 'top'})
f_sec = wb.add_format({'bold': True, 'font_size': 12, 'font_color': '1F4E79'})
f_num = wb.add_format({'num_format': '0.000', 'border': 1, 'border_color': 'D9D9D9'})
f_num5 = wb.add_format({'num_format': '0.00000', 'border': 1, 'border_color': 'D9D9D9'})
f_txt = wb.add_format({'border': 1, 'border_color': 'D9D9D9'})
f_pred = wb.add_format({'border': 1, 'border_color': 'D9D9D9', 'bg_color': 'FFF2CC'})
f_int = wb.add_format({'num_format': '0', 'border': 1, 'border_color': 'D9D9D9'})
f_link = wb.add_format({'font_color': '0563C1', 'underline': 1})

ov = wb.add_worksheet('Overview')

def xaxis(ch):
    ch.set_x_axis({'name': 'Elapsed time (days)', 'min': 0, 'max': 14, 'major_unit': 1,
                   'major_gridlines': {'visible': True, 'line': {'color': 'E7E6E6'}}})

def scatter(title, ytitle):
    ch = wb.add_chart({'type': 'scatter', 'subtype': 'straight_with_markers'})
    ch.set_title({'name': title, 'name_font': {'size': 11}})
    xaxis(ch)
    ch.set_y_axis({'name': ytitle, 'major_gridlines': {'visible': True, 'line': {'color': 'E7E6E6'}}})
    ch.set_legend({'position': 'bottom'})
    ch.show_blanks_as('span')
    ch.set_size({'width': 520, 'height': 300})
    return ch

summary = []
for b in batches:
    d = df[df.batch_id == b]
    is_test = b in test
    ws = wb.add_worksheet(b)
    ws.write(0, 0, f'{b} – ' + ('TEST batch: targets given only at day 0; yellow rows are the time points to predict'
                                if is_test else 'TRAINING batch: fully measured'), f_title)
    ws.write(1, 0, 'Source: bioreaction_remaining.csv. Units are not stated in the dataset. '
                   'Feed tables show each feed event (amount added) and the running total per feed.', f_sub)
    ws.write_url(1, 13, "internal:'Overview'!A1", f_link, string='← back to Overview')
    m = d[d[T + ['volume']].notna().any(axis=1)]
    r0 = 3
    ws.write(r0, 0, 'Measurements', f_sec)
    hdr = ['elapsed_time (days)', 'vcd', 'tcd', 'substrate', 'byproduct', 'volume', 'to_predict']
    ws.write_row(r0 + 1, 0, hdr, f_hdr)
    for i, (_, row) in enumerate(m.iterrows()):
        rr = r0 + 2 + i
        fmt = f_pred if row.to_predict else f_num
        ws.write_number(rr, 0, row.elapsed_time, fmt)
        for j, c in enumerate(T + ['volume'], start=1):
            if pd.notna(row[c]):
                ws.write_number(rr, j, row[c], f_pred if row.to_predict else f_num)
            else:
                ws.write_blank(rr, j, None, f_pred if row.to_predict else f_txt)
        ws.write(rr, 6, 'TRUE' if row.to_predict else 'FALSE', f_pred if row.to_predict else f_txt)
    m_first, m_last = r0 + 2, r0 + 1 + len(m)
    fe = d[d[F].notna().any(axis=1)].copy()
    c0 = 8
    ws.write(r0, c0, 'Feed events (amount added) and running total per feed', f_sec)
    fh = ['elapsed_time (days)'] + F + ['cum_' + f for f in F]
    ws.write_row(r0 + 1, c0, fh, f_hdr)
    cum = fe[F].fillna(0).cumsum()
    for i, (idx, row) in enumerate(fe.iterrows()):
        rr = r0 + 2 + i
        ws.write_number(rr, c0, row.elapsed_time, f_num)
        for j, c in enumerate(F, start=1):
            if pd.notna(row[c]):
                ws.write_number(rr, c0 + j, row[c], f_num5)
            else:
                ws.write_blank(rr, c0 + j, None, f_txt)
        for j, c in enumerate(F, start=1):
            ws.write_number(rr, c0 + 4 + j, cum.loc[idx, c], f_num5)
    f_first, f_last = r0 + 2, r0 + 1 + len(fe)
    ws.set_column(0, 0, 11); ws.set_column(1, 6, 10); ws.set_column(7, 7, 2)
    ws.set_column(c0, c0, 11); ws.set_column(c0 + 1, c0 + 8, 12.5)
    ws.set_row(r0 + 1, 30)
    ws.freeze_panes(r0 + 2, 0)
    def ser(ch, col, name, color, first, last, xcol, marker_only=False, y2=False):
        s = {'name': name, 'categories': [b, first, xcol, last, xcol], 'values': [b, first, col, last, col],
             'marker': {'type': 'circle', 'size': 6, 'fill': {'color': color}, 'border': {'color': color}},
             'line': {'color': color, 'width': 1.75}}
        if marker_only:
            s['line'] = {'none': True}
        if y2:
            s['y2_axis'] = True
        ch.add_series(s)
    charts = []
    ch = scatter('Cell density', 'Cell density')
    ser(ch, 1, 'VCD (viable)', COL['vcd'], m_first, m_last, 0)
    ser(ch, 2, 'TCD (total)', COL['tcd'], m_first, m_last, 0)
    charts.append(ch)
    ch = scatter('Substrate and byproduct', 'Concentration')
    ser(ch, 3, 'Substrate', COL['substrate'], m_first, m_last, 0)
    ser(ch, 4, 'Byproduct', COL['byproduct'], m_first, m_last, 0)
    charts.append(ch)
    ch = scatter('Culture volume', 'Volume')
    ser(ch, 5, 'Volume', COL['volume'], m_first, m_last, 0)
    charts.append(ch)
    if len(fe):
        ch = scatter('Cumulative feed added (all feeds)', 'Cumulative amount')
        for j, c in enumerate(F, start=1):
            ser(ch, c0 + 4 + j, 'cum ' + c, COL[c], f_first, f_last, c0)
        charts.append(ch)
        for j, c in enumerate(F, start=1):
            if fe[c].notna().sum() == 0:
                continue
            ch = scatter(f'{c}: events (dots) and running total (line)', 'Amount per event')
            ser(ch, c0 + j, f'{c} event amount', COL[c], f_first, f_last, c0, marker_only=True)
            ser(ch, c0 + 4 + j, f'{c} running total (right axis)', COL['volume'], f_first, f_last, c0, y2=True)
            ch.set_y2_axis({'name': 'Running total'})
            charts.append(ch)
    base_col = 18
    for k, ch in enumerate(charts):
        ws.insert_chart(3 + (k // 2) * 16, base_col + (k % 2) * 9, ch)
    d0 = m.iloc[0]
    summary.append([b, 'test' if is_test else 'train', int((~m.to_predict & m[T].notna().any(axis=1)).sum()),
                    int(m.to_predict.sum()), float(d.elapsed_time.max()),
                    d0.vcd, d0.tcd, d0.substrate, d0.byproduct, d0.volume,
                    *[float(d[c].sum()) for c in F], *[int(d[c].notna().sum()) for c in F],
                    len(m), m_first, m_last])

ov.write(0, 0, 'Decode the Bioreactor – one sheet per batch', f_title)
ov.write(1, 0, f'{len(batches)} batches: {len(batches) - len(test)} training (fully measured), {len(test)} test '
               '(day 0 only; to predict). Click a batch name to open its sheet. Units are not stated in the dataset.', f_sub)
oh = ['batch', 'set', 'measured time points', 'time points to predict', 'last time (days)',
      'day0 vcd', 'day0 tcd', 'day0 substrate', 'day0 byproduct', 'day0 volume',
      'total feed_substrate', 'total feed_b', 'total feed_c', 'total feed_d',
      'n feed_substrate events', 'n feed_b events', 'n feed_c events', 'n feed_d events']
ov.write_row(3, 0, oh, f_hdr)
ov.set_row(3, 45)
for i, s in enumerate(summary):
    r = 4 + i
    ov.write_url(r, 0, f"internal:'{s[0]}'!A1", f_link, string=s[0])
    ov.write(r, 1, s[1], f_pred if s[1] == 'test' else f_txt)
    for j, v in enumerate(s[2:18], start=2):
        ov.write_number(r, j, v, f_num if 4 <= j <= 13 else f_int)
ov.set_column(0, 0, 10); ov.set_column(1, 17, 11)
ov.freeze_panes(4, 1)
ov.autofilter(3, 0, 3 + len(summary), len(oh) - 1)
pos = 0
for col, var in [(1, 'vcd'), (2, 'tcd'), (3, 'substrate'), (4, 'byproduct'), (5, 'volume')]:
    ch = wb.add_chart({'type': 'scatter', 'subtype': 'straight'})
    ttl = TLAB.get(var, 'Culture volume') + (' – all training batches' if var != 'volume' else ' – all batches')
    ch.set_title({'name': ttl, 'name_font': {'size': 11}})
    xaxis(ch)
    ch.set_y_axis({'name': var, 'major_gridlines': {'visible': True, 'line': {'color': 'E7E6E6'}}})
    ch.set_legend({'none': True})
    ch.show_blanks_as('span')
    for s in summary:
        if var != 'volume' and s[1] == 'test':
            continue
        ch.add_series({'name': s[0], 'categories': [s[0], s[19], 0, s[20], 0],
                       'values': [s[0], s[19], col, s[20], col], 'line': {'width': 1}})
    ch.set_size({'width': 620, 'height': 340})
    ov.insert_chart(5 + len(summary) + 2 + (pos // 2) * 18, (pos % 2) * 9, ch)
    pos += 1
wb.close()
print('written', OUT, os.path.getsize(OUT), 'bytes;', len(batches), 'batch sheets')
