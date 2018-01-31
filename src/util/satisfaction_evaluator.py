from collections import namedtuple
import os
import openpyxl
import openpyxl.cell
import openpyxl.chart
import openpyxl.styles
import openpyxl.worksheet
import util.io


class SatisfactionEvaluator:
    ModeInfo = namedtuple('ModeInfo', ['name', 'marker'])

    solver_modes = ["unfair", "fair-linear", "fair-nonlinear"]

    mode_settings = {
        "unfair": ModeInfo('C', 'diamond'),
        "fair-linear": ModeInfo('ESA', 'x'),
        "fair-nonlinear": ModeInfo("ESD", "plus")
    }

    def __init__(self, data_dir: str, solution_dir: str) -> None:
        self.data_dir = data_dir
        self.solution_dir = solution_dir
        self.solution_data = None
        self.physicians = None
        self.results = None

        self.analysis_sheet = None
        self.charts_sheet = None

        # tables on per-directory sheets
        self.sat_first_col = 1
        self.sat_first_row = 1
        self.req_first_col = 6
        self.req_first_row = 1
        self.delta_req_first_col = 9
        self.delta_req_first_row = 1
        self.single_phys_var_first_col = 21
        self.single_phys_var_first_row = 1

        # tables on analysis sheet
        self.start_sat_table_col = 1
        self.start_sat_table_row = 1
        self.start_dir_table_col = self.start_sat_table_col + (len(self.solver_modes) * 2) + 2
        self.start_dir_table_row = 1
        self.start_req_table_col = self.start_dir_table_col + len(self.solver_modes) + 2
        self.start_req_table_row = 1
        self.start_single_phys_table_col = self.start_req_table_col + (len(self.solver_modes) * 4) + 4
        self.start_single_phys_table_row = 1

    def _find_continuous_physicians(self) -> set:
        physicians = None
        for directory in filter(lambda x: os.path.isdir(os.path.join(self.data_dir, x)),
                                sorted(os.listdir(self.data_dir))):
            data = util.io.CmplCdatReader.read(os.path.join(self.data_dir, directory, "transformed.cdat"))
            if physicians is None:
                physicians = data.get_set("J")
            else:
                physicians &= data.get_set("J")

        self.physicians = physicians
        return physicians

    def _read(self) -> dict:
        physicians = sorted(self.physicians)
        self.solution_data = dict()
        result = dict()

        for directory in filter(lambda x: os.path.isdir(os.path.join(self.solution_dir, x)),
                                sorted(os.listdir(self.solution_dir))):
            dir_result = dict()
            result[directory] = dir_result
            parameters = util.io.CmplCdatReader.read(os.path.join(self.data_dir, directory, "transformed.cdat"))
            days = parameters.get_scalar("W_max") * 7

            dir_result["total"] = {
                "g_req_on": len(parameters.get_values("g_req_on")),
                "g_req_off": len(parameters.get_values("g_req_off"))
            }

            for phys in physicians:
                dir_result[phys] = {
                    "g_req_on": parameters.count_if_exists("g_req_on", (phys, None, None, None)),
                    "g_req_off": parameters.count_if_exists("g_req_off", (phys, None, None))
                }

            dir_result["modes"] = dict()

            for solver_mode in self.solver_modes:
                mode_result = dict()
                dir_result["modes"][solver_mode] = mode_result
                solution = util.io.CmplSolutionReader.read(os.path.join(
                    self.solution_dir, directory, solver_mode + ".sol"))
                for phys in physicians:
                    req_on = dir_result[phys]["g_req_on"]
                    req_off = dir_result[phys]["g_req_off"]
                    delta_req_on = solution.count_if_exists("delta_req_on", (phys, None, None, None))
                    delta_req_off = solution.count_if_exists("delta_req_off", (phys, None, None))
                    sat = (req_on - delta_req_on + req_off - delta_req_off) / days

                    mode_result[phys] = {
                        "satisfaction": sat,
                        "delta_req_on": delta_req_on,
                        "delta_req_off": delta_req_off
                    }

                mode_result["total"] = {
                    "delta_req_on": solution.count_if_exists("delta_req_on", (None, None, None, None)),
                    "delta_req_off": solution.count_if_exists("delta_req_off", (None, None, None, None))
                }

        self.results = result
        return result

    def _write(self, filename: str):
        wb = openpyxl.Workbook()
        wb.remove_sheet(wb.get_active_sheet())

        for directory in self.results:
            sheet = wb.create_sheet(directory)

            self._write_directory_satisfaction_table(sheet, directory)
            self._write_directory_request_table(sheet, directory)
            self._write_directory_delta_request_table(sheet, directory)

        analysis_sheet = wb.create_sheet("analysis", 0)
        self.analysis_sheet = analysis_sheet
        charts_sheet = wb.create_sheet("charts", 0)
        self.charts_sheet = charts_sheet

        self._write_satisfaction_table(analysis_sheet, self.start_sat_table_col, self.start_sat_table_row)
        self._write_satisfaction_directory_table(analysis_sheet, self.start_dir_table_col, self.start_dir_table_row)
        self._write_request_table(analysis_sheet, self.start_req_table_col, self.start_req_table_row)
        self._write_single_physician_table(analysis_sheet, self.start_single_phys_table_col,
                                           self.start_single_phys_table_row)

        self._write_satisfaction_average_per_phys_chart(charts_sheet, 1, 1)
        self._write_satisfaction_variance_per_phys_chart(charts_sheet, 11, 1)
        self._write_single_physician_chart(charts_sheet, self.single_phys_var_first_col, self.single_phys_var_first_row)

        wb.save(filename)

    @staticmethod
    def _setBoldFont(cell: openpyxl.cell.Cell):
        cell.font = openpyxl.styles.Font(bold=True)

    @staticmethod
    def _setTopBorder(cell: openpyxl.cell.Cell):
        cell.border = openpyxl.styles.Border(top=openpyxl.styles.Side(border_style="thin", color="000000"))

    def _write_satisfaction_table(self, sheet: openpyxl.worksheet.Worksheet, first_col: int, first_row: int) -> None:
        cell = sheet.cell(column=first_col, row=first_row)
        cell.value = "VAR/AVG(s) by phys"
        self._setBoldFont(cell)

        self._write_physician_table(sheet, first_col, first_row + 1)

        cell = sheet.cell(row=first_row + 1 + len(self.physicians), column=first_col)
        cell.value = "average"
        self._setBoldFont(cell)
        self._setTopBorder(cell)

        cell = sheet.cell(row=first_row + 2 + len(self.physicians), column=first_col)
        cell.value = "difference"
        self._setBoldFont(cell)

        cell = sheet.cell(row=first_row + 3 + len(self.physicians), column=first_col)
        cell.value = "diff %"
        self._setBoldFont(cell)

        cell = sheet.cell(row=first_row + 4 + len(self.physicians), column=first_col)
        cell.value = "variance"
        self._setBoldFont(cell)

        cell = sheet.cell(row=first_row + 5 + len(self.physicians), column=first_col)
        cell.value = "difference"
        self._setBoldFont(cell)

        cell = sheet.cell(row=first_row + 6 + len(self.physicians), column=first_col)
        cell.value = "diff %"
        self._setBoldFont(cell)

        col_offset = 1
        for solver_mode in self.solver_modes:
            cell = sheet.cell(row=first_row, column=col_offset + first_col)
            cell.value = solver_mode
            self._setBoldFont(cell)

            cell = sheet.cell(row=first_row, column=col_offset + len(self.solver_modes) + first_col)
            cell.value = solver_mode
            self._setBoldFont(cell)

            for phys_cnt in range(len(self.physicians)):
                row = first_row + phys_cnt + 1

                cell = sheet.cell(row=row, column=first_col + col_offset)
                dirsheet_row = self.sat_first_row + phys_cnt + 1
                dirsheet_col = self.sat_first_col + col_offset
                dirsheet_cell = sheet.cell(row=dirsheet_row, column=dirsheet_col)
                cell.value = "=_xlfn.VAR.P({})".format(",".join(["'{}'!{}".format(sheet_name, dirsheet_cell.coordinate)
                                                                 for sheet_name in self.results]))

                cell = sheet.cell(row=row, column=first_col + len(self.solver_modes) + col_offset)
                dirsheet_row = self.sat_first_row + phys_cnt + 1
                dirsheet_col = self.sat_first_col + col_offset
                dirsheet_cell = sheet.cell(row=dirsheet_row, column=dirsheet_col)
                cell.value = "=AVERAGE({})".format(",".join(["'{}'!{}".format(sheet_name, dirsheet_cell.coordinate)
                                                             for sheet_name in self.results]))

            row = first_row + len(self.physicians) + 1
            background_filled = openpyxl.styles.PatternFill("solid", fgColor="d8e4bc")
            for repetition, styles, kpi_name in [(0, (background_filled, None), "APV"),
                                                 (1, (None, background_filled), "APS")]:
                column = first_col + col_offset + (repetition * len(self.solver_modes))
                first_mode_column = first_col + 1 + (repetition * len(self.solver_modes))
                cell = sheet.cell(row=row, column=column)
                cell.value = "=AVERAGE({}:{})".format(
                    sheet.cell(row=first_row, column=column).coordinate,
                    sheet.cell(row=row - 1, column=column).coordinate)
                self._setBoldFont(cell)
                self._setTopBorder(cell)
                if styles[0]:
                    cell.fill = styles[0]

                cell = sheet.cell(row=row + 1, column=column)
                cell.value = "={}-{}".format(
                    sheet.cell(row=row, column=column).coordinate,
                    sheet.cell(row=row, column=first_mode_column).coordinate)
                if styles[0]:
                    cell.fill = styles[0]

                cell = sheet.cell(row=row + 2, column=column)
                cell.value = "={}/{}*100".format(
                    sheet.cell(row=row + 1, column=column).coordinate,
                    sheet.cell(row=row, column=first_mode_column).coordinate)
                if styles[0]:
                    cell.fill = styles[0]

                cell = sheet.cell(row=row + 3, column=column)
                cell.value = "=_xlfn.VAR.P({}:{})".format(
                    sheet.cell(row=first_row, column=column).coordinate,
                    sheet.cell(row=row - 1, column=column).coordinate)
                self._setBoldFont(cell)
                if styles[1]:
                    cell.fill = styles[1]

                cell = sheet.cell(row=row + 4, column=column)
                cell.value = "={}-{}".format(
                    sheet.cell(row=row + 3, column=column).coordinate,
                    sheet.cell(row=row + 3, column=first_mode_column).coordinate)
                if styles[1]:
                    cell.fill = styles[1]

                cell = sheet.cell(row=row + 5, column=column)
                cell.value = "={}/{}*100".format(
                    sheet.cell(row=row + 4, column=column).coordinate,
                    sheet.cell(row=row + 3, column=first_mode_column).coordinate)
                if styles[1]:
                    cell.fill = styles[1]

                if (len(self.solver_modes) // 2) < col_offset <= (len(self.solver_modes) // 2) + 1:
                    cell = sheet.cell(row=row + 7, column=column)
                    cell.value = kpi_name
                    cell.fill = background_filled

            col_offset += 1

    def _write_satisfaction_directory_table(self, sheet: openpyxl.worksheet.Worksheet, first_col: int, first_row: int) \
            -> None:
        cell = sheet.cell(column=first_col, row=first_row)
        cell.value = "VAR(s) by plan"
        self._setBoldFont(cell)

        row = first_row + 1
        for directory in self.results:
            cell = sheet.cell(column=first_col, row=row)
            cell.value = directory
            self._setBoldFont(cell)
            row += 1

        cell = sheet.cell(column=first_col, row=row)
        cell.value = "average"
        self._setBoldFont(cell)
        self._setTopBorder(cell)

        cell = sheet.cell(column=first_col, row=row + 1)
        cell.value = "difference"
        self._setBoldFont(cell)

        cell = sheet.cell(column=first_col, row=row + 2)
        cell.value = "diff %"
        self._setBoldFont(cell)

        mode_cnt = 1
        for solver_mode in self.solver_modes:
            column = first_col + mode_cnt
            cell = sheet.cell(column=column, row=first_row)
            cell.value = solver_mode
            self._setBoldFont(cell)
            row = first_row + 1
            for directory in self.results:
                sheet.cell(column=column, row=row).value = "=_xlfn.VAR.P('{}'!{}:{})".format(
                    directory,
                    sheet.cell(row=self.sat_first_row + 1, column=self.sat_first_col + mode_cnt).coordinate,
                    sheet.cell(row=self.sat_first_row + len(self.physicians),
                               column=self.sat_first_col + mode_cnt).coordinate
                )
                row += 1
            cell = sheet.cell(column=column, row=row)
            cell.value = "=AVERAGE({}:{})".format(
                sheet.cell(row=first_row, column=column).coordinate,
                sheet.cell(row=row - 1, column=column).coordinate
            )
            self._setBoldFont(cell)
            self._setTopBorder(cell)

            cell = sheet.cell(row=row + 1, column=column)
            cell.value = "={}-{}".format(
                sheet.cell(row=row, column=column).coordinate,
                sheet.cell(row=row, column=first_col + 1).coordinate)

            cell = sheet.cell(row=row + 2, column=column)
            cell.value = "={}/{}*100".format(
                sheet.cell(row=row + 1, column=column).coordinate,
                sheet.cell(row=row, column=first_col + 1).coordinate
            )

            mode_cnt += 1

    def _write_physician_table(self, sheet: openpyxl.worksheet.Worksheet, first_col: int, first_row: int) -> None:
        row = first_row
        for phys in sorted(self.physicians):
            cell = sheet.cell(row=row, column=first_col)
            cell.value = phys
            self._setBoldFont(cell)

            row += 1

    def _write_directory_satisfaction_table(self, sheet: openpyxl.worksheet.Worksheet, directory: str) -> None:
        first_col = self.sat_first_col
        first_row = self.sat_first_row
        cell = sheet.cell(column=first_col, row=first_row)
        cell.value = "s"
        self._setBoldFont(cell)

        self._write_physician_table(sheet, first_col, first_row + 1)

        col_offset = 1
        for solver_mode in self.solver_modes:
            cell = sheet.cell(row=1, column=col_offset + first_col)
            cell.value = solver_mode
            self._setBoldFont(cell)
            row = first_row + 1
            for phys in sorted(self.physicians):
                sheet.cell(row=row, column=col_offset + first_col).value = \
                    self.results[directory]["modes"][solver_mode][phys]["satisfaction"]
                row += 1
            col_offset += 1

    def _write_directory_request_table(self, sheet: openpyxl.worksheet.Worksheet, directory: str) -> None:
        first_col = self.req_first_col
        first_row = self.req_first_row

        cell = sheet.cell(column=first_col, row=first_row)
        cell.value = "g_req_on"
        self._setBoldFont(cell)

        cell = sheet.cell(column=first_col + 1, row=first_row)
        cell.value = "g_req_off"
        self._setBoldFont(cell)

        for phys_cnt, phys in enumerate(sorted(self.physicians)):
            cell = sheet.cell(column=first_col, row=first_row + 1 + phys_cnt)
            cell.value = self.results[directory][phys]["g_req_on"]
            cell = sheet.cell(column=first_col + 1, row=first_row + 1 + phys_cnt)
            cell.value = self.results[directory][phys]["g_req_off"]

    def _write_directory_delta_request_table(self, sheet: openpyxl.worksheet.Worksheet, directory: str) -> None:
        first_col = self.delta_req_first_col
        first_row = self.delta_req_first_row

        for mode_cnt, mode in enumerate(self.solver_modes):
            req_on_col = first_col + (mode_cnt * 2)
            req_off_col = req_on_col + 1

            cell = sheet.cell(column=req_on_col, row=first_row)
            cell.value = mode + " delta_req_on"
            self._setBoldFont(cell)

            cell = sheet.cell(column=req_off_col, row=first_row)
            cell.value = mode + " delta_req_off"
            self._setBoldFont(cell)

            for phys_cnt, phys in enumerate(sorted(self.physicians)):
                cell = sheet.cell(column=req_on_col, row=first_row + 1 + phys_cnt)
                cell.value = self.results[directory]["modes"][mode][phys]["delta_req_on"]
                cell = sheet.cell(column=req_off_col, row=first_row + 1 + phys_cnt)
                cell.value = self.results[directory]["modes"][mode][phys]["delta_req_off"]

    def _write_request_table(self, sheet: openpyxl.worksheet.Worksheet, first_col: int, first_row: int) -> None:
        cell = sheet.cell(column=first_col, row=first_row)
        cell.value = "requests"
        self._setBoldFont(cell)

        total_col = first_col + 1
        cell = sheet.cell(column=total_col, row=first_row)
        cell.value = "total (continuous phys.)"
        self._setBoldFont(cell)

        overall_col = first_col + 2
        cell = sheet.cell(column=overall_col, row=first_row)
        cell.value = "total (all phys.)"
        self._setBoldFont(cell)

        for directory_cnt, directory in enumerate(self.results):
            dir_row = first_row + 1 + directory_cnt
            cell = sheet.cell(row=dir_row, column=first_col)
            cell.value = directory
            self._setBoldFont(cell)

            first_req_cell = sheet.cell(row=self.req_first_row + 1, column=self.req_first_col)
            last_req_cell = sheet.cell(row=self.req_first_row + len(self.physicians), column=self.req_first_col + 1)

            total_delta_cell = sheet.cell(row=dir_row, column=total_col)
            total_delta_cell.value = "=SUM('{}'!{}:{})".format(directory, first_req_cell.coordinate,
                                                               last_req_cell.coordinate)

            overall_cell = sheet.cell(row=dir_row, column=overall_col)
            overall_cell.value = self.results[directory]["total"]["g_req_on"] \
                                 + self.results[directory]["total"]["g_req_off"]

            for mode_cnt, mode in enumerate(self.solver_modes):
                first_delta_cell = sheet.cell(row=self.delta_req_first_row + 1,
                                              column=self.delta_req_first_col + (2 * mode_cnt))
                last_delta_cell = sheet.cell(row=self.delta_req_first_row + len(self.physicians),
                                             column=self.delta_req_first_col + (2 * mode_cnt) + 1)

                delta_overall_col = first_col + 3 + mode_cnt
                delta_cell = sheet.cell(row=dir_row, column=delta_overall_col)
                delta_cell.value = "=SUM('{}'!{}:{})".format(directory,
                                                             first_delta_cell.coordinate,
                                                             last_delta_cell.coordinate)

                percent_col = delta_overall_col + len(self.solver_modes)
                percent_cell = sheet.cell(row=dir_row, column=percent_col)
                percent_cell.value = "={}/{}*100".format(delta_cell.coordinate, total_delta_cell.coordinate)

                overall_delta_col = percent_col + len(self.solver_modes)
                overall_delta_cell = sheet.cell(row=dir_row, column=overall_delta_col)
                overall_delta_cell.value = self.results[directory]["modes"][mode]["total"]["delta_req_on"] \
                                           + self.results[directory]["modes"][mode]["total"]["delta_req_off"]

                overall_percent_col = overall_delta_col + len(self.solver_modes)
                overall_percent_cell = sheet.cell(row=dir_row, column=overall_percent_col)
                overall_percent_cell.value = "={}/{}*100".format(overall_delta_cell.coordinate, overall_cell.coordinate)

        row = first_row + len(self.results) + 1
        cell = sheet.cell(column=first_col, row=row)
        cell.value = "sum"
        self._setBoldFont(cell)
        self._setTopBorder(cell)

        total_sum_cell = sheet.cell(column=total_col, row=row)
        total_sum_cell.value = "=SUM({}:{})".format(sheet.cell(row=first_row + 1, column=total_col).coordinate,
                                                    sheet.cell(row=first_row + len(self.results),
                                                               column=total_col).coordinate)
        self._setBoldFont(total_sum_cell)
        self._setTopBorder(total_sum_cell)

        overall_sum_cell = sheet.cell(column=overall_col, row=row)
        overall_sum_cell.value = "=SUM({}:{})".format(sheet.cell(row=first_row + 1, column=overall_col).coordinate,
                                                      sheet.cell(row=first_row + len(self.results),
                                                                 column=overall_col).coordinate)
        self._setBoldFont(overall_sum_cell)
        self._setTopBorder(overall_sum_cell)

        for mode_cnt, mode in enumerate(self.solver_modes):
            delta_col = first_col + mode_cnt + 3
            cell = sheet.cell(column=delta_col, row=first_row)
            cell.value = "delta (continuous phys.) " + mode
            self._setBoldFont(cell)

            vio_col = delta_col + len(self.solver_modes)
            cell = sheet.cell(column=vio_col, row=first_row)
            cell.value = "%vio (continuous phys.) " + mode
            self._setBoldFont(cell)

            delta_overall_col = vio_col + len(self.solver_modes)
            cell = sheet.cell(column=delta_overall_col, row=first_row)
            cell.value = "delta (all phys.) " + mode
            self._setBoldFont(cell)

            vio_overall_col = delta_overall_col + len(self.solver_modes)
            cell = sheet.cell(column=vio_overall_col, row=first_row)
            cell.value = "%vio (all phys.) " + mode
            self._setBoldFont(cell)

            total_delta_cell = sheet.cell(column=delta_col, row=row)
            total_delta_cell.value = "=SUM({}:{})".format(sheet.cell(row=first_row + 1,
                                                                     column=delta_col).coordinate,
                                                          sheet.cell(row=first_row + len(self.results),
                                                                     column=delta_col).coordinate)
            self._setBoldFont(total_delta_cell)
            self._setTopBorder(total_delta_cell)

            percentage_cell = sheet.cell(column=vio_col, row=row)
            percentage_cell.value = "={}/{}*100".format(total_delta_cell.coordinate, total_sum_cell.coordinate)
            self._setBoldFont(percentage_cell)
            self._setTopBorder(percentage_cell)

            total_delta_overall_cell = sheet.cell(column=delta_overall_col, row=row)
            total_delta_overall_cell.value = "=SUM({}:{})".format(sheet.cell(row=first_row + 1,
                                                                             column=delta_overall_col).coordinate,
                                                                  sheet.cell(row=first_row + len(self.results),
                                                                             column=delta_overall_col).coordinate)
            self._setBoldFont(total_delta_overall_cell)
            self._setTopBorder(total_delta_overall_cell)

            total_vio_overall_cell = sheet.cell(column=vio_overall_col, row=row)
            total_vio_overall_cell.value = "={}/{}*100".format(total_delta_overall_cell.coordinate,
                                                               overall_sum_cell.coordinate)
            self._setBoldFont(total_vio_overall_cell)
            self._setTopBorder(total_vio_overall_cell)

    def _write_satisfaction_average_per_phys_chart(self, sheet: openpyxl.worksheet.Worksheet, first_col: int,
                                                   first_row: int) -> None:
        anchor = sheet.cell(column=first_col, row=first_row)
        cell = sheet.cell(column=first_col, row=first_row)
        cell.value = "average satisfaction per physician"

        chart = openpyxl.chart.LineChart()
        chart.style = 1
        chart.height = 9
        chart.width = 15
        chart.legend.position = "b"
        chart.x_axis.title = "physician"
        chart.y_axis.title = "average satisfaction"

        xref = openpyxl.chart.Reference(self.analysis_sheet, min_col=self.start_sat_table_col,
                                        max_col=self.start_sat_table_col,
                                        min_row=self.start_sat_table_row + 1,
                                        max_row=self.start_sat_table_row + len(self.physicians))
        chart.set_categories(xref)
        for mode_cnt, mode in enumerate(self.solver_modes):
            data_column = self.start_sat_table_col + 1 + len(self.solver_modes) + mode_cnt
            ref = openpyxl.chart.Reference(self.analysis_sheet, min_col=data_column, max_col=data_column,
                                           min_row=self.start_sat_table_row + 1,
                                           max_row=self.start_sat_table_row + len(self.physicians))
            series = openpyxl.chart.Series(ref, title=self.mode_settings[mode].name)
            series.marker.symbol = self.mode_settings[mode].marker
            series.graphicalProperties.line.noFill = True
            chart.append(series)

        sheet.add_chart(chart, anchor=anchor.offset(column=1).coordinate)

    def _write_satisfaction_variance_per_phys_chart(self, sheet: openpyxl.worksheet.Worksheet, first_col: int,
                                                    first_row: int) -> None:
        cell = sheet.cell(column=first_col, row=first_row)
        cell.value = "variance per physician"

        chart = openpyxl.chart.LineChart()
        chart.style = 1
        chart.height = 9
        chart.width = 15
        chart.legend.position = "b"
        chart.x_axis.title = "physician"
        chart.y_axis.title = "variance"

        xref = openpyxl.chart.Reference(self.analysis_sheet, min_col=self.start_sat_table_col,
                                        max_col=self.start_sat_table_col,
                                        min_row=self.start_sat_table_row + 1,
                                        max_row=self.start_sat_table_row + len(self.physicians))
        chart.set_categories(xref)
        for mode_cnt, mode in enumerate(self.solver_modes):
            data_column = self.start_sat_table_col + 1 + mode_cnt
            ref = openpyxl.chart.Reference(self.analysis_sheet, min_col=data_column, max_col=data_column,
                                           min_row=self.start_sat_table_row + 1,
                                           max_row=self.start_sat_table_row + len(self.physicians))
            series = openpyxl.chart.Series(ref, title=self.mode_settings[mode].name)
            series.marker.symbol = self.mode_settings[mode].marker
            series.graphicalProperties.line.noFill = True
            chart.append(series)

        sheet.add_chart(chart, anchor=cell.offset(column=1).coordinate)

    def _write_single_physician_table(self, sheet: openpyxl.worksheet.Worksheet, first_col: int, first_row: int) \
            -> None:
        physician_data_cell = self.charts_sheet.cell(column=self.single_phys_var_first_col,
                                                     row=self.single_phys_var_first_row).offset(column=1)

        for mode_cnt, mode in enumerate(self.solver_modes):
            cell = sheet.cell(column=first_col + 1 + mode_cnt, row=first_row)
            cell.value = mode
            self._setBoldFont(cell)

        for dir_cnt, directory in enumerate(sorted(self.results)):
            row = first_row + 1 + dir_cnt

            header_cell = sheet.cell(row=row, column=first_col)
            header_cell.value = directory
            self._setBoldFont(header_cell)

            for mode_cnt, mode in enumerate(self.solver_modes):
                col = first_col + 1 + mode_cnt
                mode_col_letter = chr(ord('B') + mode_cnt)

                value_cell = sheet.cell(row=row, column=col)
                value_cell.value = "=INDIRECT(\"'\" & {} & \"'!{}\" & ('{}'!{} + 1))".format(
                    header_cell.coordinate, mode_col_letter, self.charts_sheet.title, physician_data_cell.coordinate)

    def _write_single_physician_chart(self, sheet: openpyxl.worksheet.Worksheet, first_col: int, first_row: int) \
            -> None:
        cell = sheet.cell(column=first_col, row=first_row)
        cell.value = "satisfaction for phys"

        cell.offset(column=1).value = 1

        chart = self._create_default_chart()
        chart.x_axis_title = "physician"
        chart.y_axis_title = "satisfaction"

        xref = openpyxl.chart.Reference(self.analysis_sheet, min_col=self.start_single_phys_table_col,
                                        max_col=self.start_single_phys_table_col,
                                        min_row=self.start_single_phys_table_row + 1,
                                        max_row=self.start_single_phys_table_row + len(self.results))
        chart.set_categories(xref)
        for mode_cnt, mode in enumerate(self.solver_modes):
            data_column = self.start_single_phys_table_col + 1 + mode_cnt
            ref = openpyxl.chart.Reference(self.analysis_sheet, min_col=data_column, max_col=data_column,
                                           min_row=self.start_single_phys_table_row + 1,
                                           max_row=self.start_single_phys_table_row + len(self.results))
            series = openpyxl.chart.Series(ref, title=self.mode_settings[mode].name)
            chart.append(series)

        sheet.add_chart(chart, anchor=cell.offset(column=1, row=1).coordinate)

    @staticmethod
    def _create_default_chart():
        chart = openpyxl.chart.LineChart()
        chart.style = 1
        chart.height = 9
        chart.height = 15
        chart.legend.position = "b"
        return chart

    def evaluate_to_file(self, filename: str):
        self._find_continuous_physicians()
        self._read()
        self._write(filename)
