import pandas as pd
import logging
from io import BytesIO
from utils import METRICS, ANALYSIS_VALUES

class ExcelGenerator:
    """Handles Excel report generation with multiple sheets and analysis"""
    
    def __init__(self, cost_params, custom_weeks=0):
        self.cost_per_stop, self.cost_per_route, self.cost_per_mile = cost_params
        self.custom_weeks = custom_weeks
        self.logger = logging.getLogger(__name__)
    
    def create_report(self, df1, df2, df3):
        """Create comprehensive Excel report with all sheets"""
        output_buffer = BytesIO()
        
        try:
            with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
                # Create all sheets
                self._create_analysis_sheet(writer, df1)
                self._create_pivot_table(writer, df1)
                df3.to_excel(writer, sheet_name='Dispatch Summaries Raw', index=False)
                df2.to_excel(writer, sheet_name='MDT Raw', index=False)
                df1.to_excel(writer, sheet_name='Totals Raw', index=False)
                self._create_mdt_chart(writer, df2)
            
            self.logger.info("Excel report generated successfully")
            
        except Exception as e:
            self.logger.error(f"Error creating Excel report: {e}")
            raise
        
        output_buffer.seek(0)
        return output_buffer
    
    def _create_analysis_sheet(self, writer, df1):
        """Create the main analysis sheet with metrics and cost calculations"""
        try:
            workbook = writer.book
            sheet = workbook.add_worksheet('Analysis')
            sheet.set_tab_color('#9BBB59')
            
            # Define formats
            title_format = workbook.add_format({
                'align': 'center', 'valign': 'vcenter', 'bg_color': '#FFC000', 
                'bold': True, 'font_size': 16
            })
            header_format = workbook.add_format({
                'align': 'center', 'valign': 'vcenter', 'bg_color': '#DCE6F1', 
                'bold': True, 'font_size': 11, 'border': 1
            })
            value_format = workbook.add_format({
                'font_size': 11, 'border': 1, 'num_format': '#,##0.00;(#,##0.00)'
            })
            annual_format = workbook.add_format({
                'font_size': 11, 'num_format': '#,##0.00;(#,##0.00)'
            })
            cost_format = workbook.add_format({
                'font_size': 11, 'num_format': '$#,##0.00;($#,##0.00)'
            })
            percent_format = workbook.add_format({
                'font_size': 11, 'border': 1, 'num_format': '0.00%'
            })
            
            # Main analysis section
            sheet.merge_range('B2:F2', 'Analysis Summary', title_format)
            
            # Headers
            headers = ['Values', 'Baseline', 'Scenario', 'Delta', 'Delta %']
            for i, header in enumerate(headers, start=1):
                sheet.write(2, i, header, header_format)
            
            # Analysis values
            for i, val in enumerate(ANALYSIS_VALUES, start=3):
                sheet.write(i, 1, val, value_format)
                sheet.write(i, 2, 0, value_format)  # Baseline (to be filled)
                sheet.write(i, 3, 0, value_format)  # Scenario (to be filled)
                sheet.write_formula(i, 4, f'=D{i+1}-C{i+1}', value_format)  # Delta
                sheet.write_formula(i, 5, f'=E{i+1}/C{i+1}', percent_format)  # Delta %
            
            # Special formulas for calculated metrics
            sheet.write_formula('C11', '=C6/C8', value_format)  # CPT baseline
            sheet.write_formula('D11', '=D6/D8', value_format)  # CPT scenario
            sheet.write_formula('C12', '=C10/C8', value_format)  # LOH baseline
            sheet.write_formula('D12', '=D10/D8', value_format)  # LOH scenario
            
            # Annualization section
            modeled_weeks = self._calculate_modeled_weeks(df1)
            annual_factor = round(52 / modeled_weeks)
            
            annual_labels = ['Stops Annually', 'Routes Annually', 'Distance Annually']
            annual_formulas = [
                f'=E9*{annual_factor}',
                f'=E8*{annual_factor}',
                f'=E10*{annual_factor}'
            ]
            
            for i, (label, formula) in enumerate(zip(annual_labels, annual_formulas), start=14):
                sheet.write(i, 1, label, workbook.add_format({'font_size': 11, 'bold': True}))
                sheet.write_formula(i, 2, formula, annual_format)
            
            # Cost impact section
            sheet.write(13, 3, "Cost", workbook.add_format({'font_size': 11, 'bold': True}))
            sheet.write(13, 4, "Annualized Routing Impact", workbook.add_format({'font_size': 11, 'bold': True}))
            
            costs = [self.cost_per_stop, self.cost_per_route, self.cost_per_mile]
            for i, cost in enumerate(costs, start=14):
                sheet.write(i, 3, cost, cost_format)
                sheet.write_formula(i, 4, f'=D{i+1}*C{i+1}', cost_format)
            
            # Total cost impact
            sheet.write_formula(17, 4, '=SUM(E15:E17)', workbook.add_format({
                'font_size': 11, 'bold': True, 'num_format': '$#,##0.00;($#,##0.00)'
            }))
            
            # Disclaimer
            sheet.write(18, 4, 
                "This cost represents the routing efficiency impacts only and does not currently include asset or driver impacts",
                workbook.add_format({'font_size': 11, 'bold': True})
            )
            
            # Summary section
            sheet.write('H2', 'Summary', workbook.add_format({'bold': True, 'font_size': 26}))
            for row in range(2, 7):
                sheet.write(row, 7, 'X', workbook.add_format({'font_size': 14}))
            
            # Set column widths
            sheet.set_column(1, 4, 17)
            
        except Exception as e:
            self.logger.error(f"Error in create_analysis_sheet: {e}")
            raise
    
    def _create_pivot_table(self, writer, df):
        """Create pivot table summary sheet"""
        try:
            if df.empty:
                return
                
            pivot = pd.pivot_table(
                df, 
                values=METRICS, 
                index='DC', 
                columns='Simulation', 
                aggfunc='sum',
                fill_value=0
            )
            
            workbook = writer.book
            sheet = workbook.add_worksheet('Summary')
            number_format = workbook.add_format({'num_format': '#,##0.00;(#,##0.00)'})
            
            # Write pivot table structure
            sheet.write('A2', 'Row Labels')
            
            row = 3
            max_row = row
            
            # Write row labels and data
            for dc in pivot.index:
                sheet.write(row, 0, dc)
                row += 1
                for val in METRICS:
                    sheet.write(row, 0, f'Sum of {val}')
                    row += 1
                    max_row = max(max_row, row)
            
            # Write column headers and values
            col = 1
            for sim in pivot.columns.levels[1] if hasattr(pivot.columns, 'levels') else pivot.columns:
                sheet.write(1, col, sim)
                row = 3
                for dc in pivot.index:
                    row += 1
                    for val in METRICS:
                        value = pivot.loc[dc, (val, sim)] if hasattr(pivot.columns, 'levels') else pivot.loc[dc, sim]
                        sheet.write(row, col, value, number_format)
                        row += 1
                        max_row = max(max_row, row)
                col += 1
            
            # Add difference column
            if col > 2:  # Only if we have at least 2 simulation columns
                sheet.write(1, col, 'Difference')
                for row in range(4, max_row):
                    sheet.write_formula(row, col, f'=C{row+1}-B{row+1}')
            
            # Add grand totals
            for i, val in enumerate(METRICS, start=row + 1):
                max_row += 1
                sheet.write(max_row, 0, f'Grand Total {val}')
                for c in range(1, min(4, col + 1)):
                    col_letter = chr(65 + c)
                    sheet.write_formula(
                        max_row, c, 
                        f'SUMIF(A4:A{max_row - 1}, "Sum of {val}", {col_letter}4:{col_letter}{max_row - 1})', 
                        number_format
                    )
            
            sheet.set_column(0, 3, 17)
            
        except Exception as e:
            self.logger.error(f"Error in create_pivot_table: {e}")
    
    def _create_mdt_chart(self, writer, df):
        """Create MDT analysis sheet with time-based charts"""
        try:
            if df.empty or 'Time Range' not in df.columns:
                return
                
            # Create pivot table for time analysis
            pivot = pd.pivot_table(
                df,
                values='Trailer',
                index='Time Range',
                columns='Simulation',
                aggfunc='count',
                fill_value=0
            )
            
            # Add Grand Total column
            pivot['Grand Total'] = pivot.sum(axis=1)
            
            # Create worksheet
            workbook = writer.book
            sheet = workbook.add_worksheet('MDT Analysis')
            bold_format = workbook.add_format({'bold': True})
            header_format = workbook.add_format({'bold': True, 'bg_color': '#DCE6F1', 'border': 1})
            number_format = workbook.add_format({'num_format': '#,##0', 'border': 1})
            
            # Write headers
            sheet.write('A2', 'Count of Trailer', bold_format)
            sheet.write('B2', 'Column Labels', bold_format)
            sheet.write('A3', 'Row Labels', header_format)
            
            for col_num, col_name in enumerate(pivot.columns, start=1):
                sheet.write(2, col_num, col_name, header_format)
            
            # Write data
            for row_num, (time_range, row_data) in enumerate(pivot.iterrows(), start=3):
                sheet.write(row_num, 0, time_range, header_format)
                for col_num, value in enumerate(row_data, start=1):
                    sheet.write(row_num, col_num, value, number_format)
            
            # Add chart if we have simulation data
            if 'Simulation' in df.columns and df['Simulation'].nunique() > 0:
                chart = workbook.add_chart({'type': 'column', 'subtype': 'clustered'})
                
                simulations = df['Simulation'].unique()
                for i, sim in enumerate(simulations):
                    if i < len(pivot.columns) - 1:  # Exclude Grand Total column
                        chart.add_series({
                            'name': ['MDT Analysis', 2, i + 1],
                            'categories': ['MDT Analysis', 3, 0, 3 + len(pivot) - 1, 0],
                            'values': ['MDT Analysis', 3, i + 1, 3 + len(pivot) - 1, i + 1],
                        })
                
                chart.set_x_axis({'name': 'Time Range'})
                chart.set_y_axis({'name': 'Count of Trailers'})
                chart.set_title({'name': 'Trailer Distribution by Time Range'})
                sheet.insert_chart('G2', chart)
            
        except Exception as e:
            self.logger.error(f"Error in create_mdt_chart: {e}")
    
    def _calculate_modeled_weeks(self, df):
        """Calculate the number of weeks represented in the data"""
        if self.custom_weeks > 0:
            return self.custom_weeks
        
        try:
            if 'Date' in df.columns and not df['Date'].dropna().empty:
                unique_dates = pd.to_datetime(df['Date'].dropna().unique())
                return max(1, round(len(unique_dates) / 7))
        except Exception as e:
            self.logger.warning(f"Could not calculate modeled weeks: {e}")
        
        return 1  # Default to 1 week
