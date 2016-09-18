import sys
from contextlib import contextmanager
from math import exp, log

import win32com.client as com
from win32com.client import constants as C

YEAR_DAYS = 365.25

class HyperbolicDecline:
    def __init__(self, qi, Di, b):
        self.qi = qi
        self.Di = Di
        self.b = b

    def rate(self, t):
        if self.b == 0:
            return self.qi * exp(-self.Di * t)
        elif self.b == 1:
            return self.qi / (1 + self.Di * t)
        else:
            return self.qi * (1 + self.b * self.Di * t) ** (-1 / self.b)

    def cumulative(self, t):
        qi_yearly = self.qi * YEAR_DAYS
        if self.Di == 0:
            return qi_yearly * t
        elif self.b == 0:
            return qi_yearly / self.Di * (1 - exp(-self.Di * t))
        elif self.b == 1:
            return qi_yearly / self.Di * log(1 + self.Di * t)
        else:
            return (qi_yearly / ((1 - self.b) * self.Di) *
              (1 - (1 + self.b * self.Di * t) ** (1 - (1 / self.b))))

class DuongDecline:
    def __init__(self, q1, a, m):
        self.q1 = q1
        self.a = a
        self.m = m

    def rate(self, t):
        if t < 1 / YEAR_DAYS:
            return self.rate(1 / YEAR_DAYS)
        return self.q1 * self._t_ab(t)

    def cumulative(self, t):
        q_yearly = self.rate(t) * YEAR_DAYS
        if t < 1 / YEAR_DAYS:
            return self.cumulative(1 / YEAR_DAYS)
        return q_yearly / (self.a * t ** -self.m)

    def _t_ab(self, t):
        return (t ** -self.m) * exp(self.a * (t ** (1 - self.m) - 1) /
                (1 - self.m))

def volume(decline, t1, t2):
    return decline.cumulative(t2) - decline.cumulative(t1)

@contextmanager
def guard(obj, guard_fn, *args):
    try:
        yield obj
    finally:
        getattr(obj, guard_fn)(*args)

def main(argv):
    EXCEL_FILES = 'Excel Files (*.xlsx; *.xlsm; *.xls),*.xlsx;*.xlsm;*.xls'

    if len(argv) >= 2:
        print('Usage: {} [forecast-months]'.format(argv[0]), file=sys.stderr)
        return 0

    if len(argv) == 2:
        try:
            forecast_months = int(argv[1])
        except ValueError:
            print('Invalid forecast-months: {}'.format(argv[1]),
                    file=sys.stderr)
            return 0
    else:
        forecast_months = 48

    with guard(com.gencache.EnsureDispatch('Excel.Application'), 'Quit') as xl:
        source_wb_fn = xl.GetOpenFilename(FileFilter=EXCEL_FILES,
                Title='Open Decline Workbook')
        if source_wb_fn == False:
            return 0

        with guard(xl.Workbooks.Open(source_wb_fn),
                'Close', False) as source_wb:
            well_declines = extract_declines(source_wb)
            well_forecasts = ((name, monthly_forecast(decl, forecast_months))
                    for (name, decl) in well_declines)

            with guard(xl.Workbooks.Add(), 'Close', False) as dest_wb:
                initial_sheets = dest_wb.Sheets.Count

                for (name, fc) in well_forecasts:
                    add_well_sheet(dest_wb, name, fc)

                xl.DisplayAlerts = False
                for _ in range(initial_sheets):
                    dest_wb.Sheets(1).Delete()
                xl.DisplayAlerts = True

                dest_wb_fn = xl.GetSaveAsFilename(
                        InitialFilename='monthly_production.xlsx',
                        FileFilter=EXCEL_FILES,
                        Title='Save Monthly Production Workbook')
                if dest_wb_fn != False:
                    dest_wb.SaveAs(dest_wb_fn)

def extract_declines(wb):
    DATA_BEGIN_ROW = 2
    WELL_NAME_COLUMN = 1
    DECLINE_TYPE_COLUMN = 2
    BEGIN_PARAMS_COLUMN = 3

    TYPES = {
        'Hyperbolic': HyperbolicDecline,
        'Duong': DuongDecline
    }

    sheet = wb.Sheets(1)

    row = DATA_BEGIN_ROW
    while sheet.Cells(row, WELL_NAME_COLUMN).Value:
        wellname = sheet.Cells(row, WELL_NAME_COLUMN).Value
        decline_type = sheet.Cells(row, DECLINE_TYPE_COLUMN).Value
        p1 = sheet.Cells(row, BEGIN_PARAMS_COLUMN).Value
        p2 = sheet.Cells(row, BEGIN_PARAMS_COLUMN + 1).Value
        p3 = sheet.Cells(row, BEGIN_PARAMS_COLUMN + 2).Value

        yield (wellname, TYPES[decline_type](p1, p2, p3))

        row += 1

def monthly_forecast(decline, months):
    return (volume(decline, t / 12, (t + 1) / 12) for t in range(months))

def add_well_sheet(wb, name, fc):
    sheet = wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))
    sheet.Name = name

    sheet.Range("A1").Value = name
    sheet.Range("A1").Font.Bold = True
    sheet.Range("A1:B1").Merge
    sheet.Range("A2").Value = "Month"
    sheet.Range("A2").Font.Bold = True
    sheet.Range("B2").Value = "Volume"
    sheet.Range("B2").Font.Bold = True

    fc = list(fc)
    months = len(fc)

    sheet.Range(
        sheet.Cells(3, 1),
        sheet.Cells(3 + months - 1, 1)
    ).Value = [[m] for m in range(1, months + 1)]

    sheet.Range(
        sheet.Cells(3, 2),
        sheet.Cells(3 + months - 1, 2)
    ).Value = [[vol] for vol in fc]

    add_graph(sheet, months)

def add_graph(sheet, months):
    graph = sheet.Parent.Charts.Add().Location(Where=C.xlLocationAsObject,
            Name=sheet.Name)

    graph.ChartType = C.xlXYScatterLinesNoMarkers

    graph.SetSourceData(sheet.Range(
        sheet.Cells(3, 1),
        sheet.Cells(3 + months - 1, 2)
    ))

    graph.Axes(C.xlValue).ScaleType = C.xlScaleLogarithmic
    graph.HasLegend = False

    graph.HasTitle = True
    graph.ChartTitle.Text = sheet.Name

if __name__ == '__main__':
    sys.exit(main(sys.argv))
