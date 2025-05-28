"""
Excel Pivot Charts Automation Tools
Comprehensive module for creating, managing, and automating Excel pivot charts
using xlwings for MCP server integration.
"""

import xlwings as xw
from typing import Dict, List, Optional, Union, Tuple, Any
import logging
from enum import Enum

# Setup logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class ChartType(Enum):
    """Excel chart types mapping"""
    COLUMN = -4100
    BAR = -4101
    LINE = -4102
    PIE = -4103
    AREA = -4098
    DOUGHNUT = -4120
    RADAR = -4151
    COMBO = 99
    COLUMN_STACKED = 52
    COLUMN_STACKED_100 = 53
    BAR_STACKED = 58
    BAR_STACKED_100 = 59
    LINE_MARKERS = 65
    LINE_STACKED = 66
    AREA_STACKED = 76
    SCATTER = -4169

class LegendPosition(Enum):
    """Legend position constants"""
    BOTTOM = -4107
    CORNER = 2
    TOP = -4160
    RIGHT = -4152
    LEFT = -4131

class AxisType(Enum):
    """Axis type constants"""
    CATEGORY = 1
    VALUE = 2
    SERIES = 3

class ExcelChartsCore:
    """Core class for Excel pivot charts automation"""
    
    def __init__(self, workbook_path: Optional[str] = None):
        """Initialize with optional workbook path"""
        self.app = None
        self.wb = None
        self.ws = None
        if workbook_path:
            self.open_workbook(workbook_path)
    
    def _get_or_create_app(self) -> xw.App:
        """Get existing Excel app or create new one"""
        try:
            if not self.app:
                self.app = xw.App(visible=True, add_book=False)
            return self.app
        except Exception as e:
            logger.error(f"Error getting Excel app: {e}")
            raise
    
    def open_workbook(self, path: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Open workbook and optionally select sheet"""
        try:
            self.app = self._get_or_create_app()
            self.wb = self.app.books.open(path)
            
            if sheet_name:
                self.ws = self.wb.sheets[sheet_name]
            else:
                self.ws = self.wb.sheets.active
                
            return {
                "status": "success",
                "workbook": self.wb.name,
                "active_sheet": self.ws.name,
                "sheets": [sheet.name for sheet in self.wb.sheets]
            }
        except Exception as e:
            logger.error(f"Error opening workbook: {e}")
            return {"status": "error", "message": str(e)}
    
    def create_pivot_table(self, data_range: str, pivot_table_name: str, 
                          destination: str, row_fields: List[str] = None,
                          column_fields: List[str] = None, value_fields: List[str] = None) -> Dict[str, Any]:
        """Create a pivot table from data range"""
        try:
            if not self.wb:
                raise ValueError("No workbook opened")
            
            # Get data range
            source_range = self.ws.range(data_range)
            
            # Create pivot cache
            pivot_cache = self.wb.api.PivotCaches().Create(
                SourceType=1,  # xlDatabase
                SourceData=source_range.api
            )
            
            # Create pivot table
            destination_sheet = self.ws if '!' not in destination else self.wb.sheets[destination.split('!')[0]]
            dest_range = destination_sheet.range(destination.split('!')[-1]) if '!' in destination else self.ws.range(destination)
            
            pivot_table = pivot_cache.CreatePivotTable(
                TableDestination=dest_range.api,
                TableName=pivot_table_name
            )
            
            # Configure fields
            if row_fields:
                for field in row_fields:
                    pivot_table.PivotFields(field).Orientation = 1  # xlRowField
            
            if column_fields:
                for field in column_fields:
                    pivot_table.PivotFields(field).Orientation = 2  # xlColumnField
            
            if value_fields:
                for field in value_fields:
                    pivot_table.PivotFields(field).Orientation = 4  # xlDataField
            
            return {
                "status": "success",
                "pivot_table_name": pivot_table_name,
                "location": destination
            }
            
        except Exception as e:
            logger.error(f"Error creating pivot table: {e}")
            return {"status": "error", "message": str(e)}
    
    def create_pivot_chart_from_table(self, pivot_table_name: str, chart_type: str = "COLUMN",
                                    chart_title: Optional[str] = None, position: Tuple[int, int] = (100, 100)) -> Dict[str, Any]:
        """Create pivot chart from existing pivot table"""
        try:
            if not self.wb:
                raise ValueError("No workbook opened")
            
            # Find pivot table
            pivot_table = None
            for sheet in self.wb.sheets:
                try:
                    pivot_table = sheet.api.PivotTables(pivot_table_name)
                    self.ws = sheet  # Set active sheet to where pivot table is
                    break
                except:
                    continue
            
            if not pivot_table:
                raise ValueError(f"Pivot table '{pivot_table_name}' not found")
            
            # Create pivot chart
            chart_type_value = getattr(ChartType, chart_type.upper()).value
            chart = self.ws.api.ChartObjects().Add(position[0], position[1], 400, 300).Chart
            
            # Set chart source to pivot table
            chart.SetSourceData(pivot_table.TableRange1)
            chart.ChartType = chart_type_value
            
            # Set title if provided
            if chart_title:
                chart.HasTitle = True
                chart.ChartTitle.Text = chart_title
            
            return {
                "status": "success",
                "chart_name": chart.Name,
                "chart_type": chart_type,
                "pivot_table": pivot_table_name
            }
            
        except Exception as e:
            logger.error(f"Error creating pivot chart: {e}")
            return {"status": "error", "message": str(e)}
    
    def create_chart_from_range(self, data_range: str, chart_type: str = "COLUMN",
                              chart_title: Optional[str] = None, position: Tuple[int, int] = (100, 100)) -> Dict[str, Any]:
        """Create chart directly from data range"""
        try:
            if not self.wb:
                raise ValueError("No workbook opened")
            
            # Get data range
            source_range = self.ws.range(data_range)
            
            # Create chart
            chart_type_value = getattr(ChartType, chart_type.upper()).value
            chart_obj = self.ws.api.ChartObjects().Add(position[0], position[1], 400, 300)
            chart = chart_obj.Chart
            
            # Set data source
            chart.SetSourceData(source_range.api)
            chart.ChartType = chart_type_value
            
            # Set title if provided
            if chart_title:
                chart.HasTitle = True
                chart.ChartTitle.Text = chart_title
            
            return {
                "status": "success",
                "chart_name": chart.Name,
                "chart_type": chart_type,
                "data_range": data_range
            }
            
        except Exception as e:
            logger.error(f"Error creating chart from range: {e}")
            return {"status": "error", "message": str(e)}
    
    def update_chart_data_source(self, chart_name: str, new_data_range: str) -> Dict[str, Any]:
        """Switch chart data source to new range"""
        try:
            chart = self._find_chart(chart_name)
            if not chart:
                raise ValueError(f"Chart '{chart_name}' not found")
            
            new_range = self.ws.range(new_data_range)
            chart.SetSourceData(new_range.api)
            
            return {
                "status": "success",
                "chart_name": chart_name,
                "new_data_range": new_data_range
            }
            
        except Exception as e:
            logger.error(f"Error updating chart data source: {e}")
            return {"status": "error", "message": str(e)}
    
    def refresh_pivot_data(self, pivot_table_name: Optional[str] = None) -> Dict[str, Any]:
        """Refresh pivot table and associated charts"""
        try:
            if pivot_table_name:
                # Refresh specific pivot table
                pivot_table = None
                for sheet in self.wb.sheets:
                    try:
                        pivot_table = sheet.api.PivotTables(pivot_table_name)
                        pivot_table.RefreshTable()
                        break
                    except:
                        continue
                
                if not pivot_table:
                    raise ValueError(f"Pivot table '{pivot_table_name}' not found")
            else:
                # Refresh all pivot tables and charts
                self.wb.api.RefreshAll()
            
            return {"status": "success", "message": "Data refreshed successfully"}
            
        except Exception as e:
            logger.error(f"Error refreshing pivot data: {e}")
            return {"status": "error", "message": str(e)}
    
    def set_chart_title(self, chart_name: str, title: str, show_title: bool = True) -> Dict[str, Any]:
        """Set or update chart title"""
        try:
            chart = self._find_chart(chart_name)
            if not chart:
                raise ValueError(f"Chart '{chart_name}' not found")
            
            chart.HasTitle = show_title
            if show_title:
                chart.ChartTitle.Text = title
            
            return {"status": "success", "chart_name": chart_name, "title": title}
            
        except Exception as e:
            logger.error(f"Error setting chart title: {e}")
            return {"status": "error", "message": str(e)}
    
    def set_axis_title(self, chart_name: str, axis_type: str, title: str, show_title: bool = True) -> Dict[str, Any]:
        """Set axis title (X or Y axis)"""
        try:
            chart = self._find_chart(chart_name)
            if not chart:
                raise ValueError(f"Chart '{chart_name}' not found")
            
            axis_num = 1 if axis_type.upper() == 'X' else 2  # 1=Category(X), 2=Value(Y)
            axis = chart.Axes(axis_num)
            
            axis.HasTitle = show_title
            if show_title:
                axis.AxisTitle.Text = title
            
            return {"status": "success", "chart_name": chart_name, "axis": axis_type, "title": title}
            
        except Exception as e:
            logger.error(f"Error setting axis title: {e}")
            return {"status": "error", "message": str(e)}
    
    def toggle_data_labels(self, chart_name: str, show_labels: bool = True, series_index: int = 1) -> Dict[str, Any]:
        """Toggle data labels on chart series"""
        try:
            chart = self._find_chart(chart_name)
            if not chart:
                raise ValueError(f"Chart '{chart_name}' not found")
            
            series = chart.SeriesCollection(series_index)
            series.HasDataLabels = show_labels
            
            return {"status": "success", "chart_name": chart_name, "data_labels": show_labels}
            
        except Exception as e:
            logger.error(f"Error toggling data labels: {e}")
            return {"status": "error", "message": str(e)}
    
    def toggle_gridlines(self, chart_name: str, axis_type: str, major: bool = True, show: bool = True) -> Dict[str, Any]:
        """Toggle major/minor gridlines"""
        try:
            chart = self._find_chart(chart_name)
            if not chart:
                raise ValueError(f"Chart '{chart_name}' not found")
            
            axis_num = 1 if axis_type.upper() == 'X' else 2
            axis = chart.Axes(axis_num)
            
            if major:
                axis.HasMajorGridlines = show
            else:
                axis.HasMinorGridlines = show
            
            return {"status": "success", "chart_name": chart_name, "gridlines": f"{axis_type} {'major' if major else 'minor'}"}
            
        except Exception as e:
            logger.error(f"Error toggling gridlines: {e}")
            return {"status": "error", "message": str(e)}
    
    def set_legend_properties(self, chart_name: str, show_legend: bool = True, position: str = "RIGHT") -> Dict[str, Any]:
        """Configure chart legend"""
        try:
            chart = self._find_chart(chart_name)
            if not chart:
                raise ValueError(f"Chart '{chart_name}' not found")
            
            chart.HasLegend = show_legend
            if show_legend:
                legend_pos = getattr(LegendPosition, position.upper()).value
                chart.Legend.Position = legend_pos
            
            return {"status": "success", "chart_name": chart_name, "legend": show_legend, "position": position}
            
        except Exception as e:
            logger.error(f"Error setting legend: {e}")
            return {"status": "error", "message": str(e)}
    
    def apply_chart_layout(self, chart_name: str, layout_id: int) -> Dict[str, Any]:
        """Apply predefined chart layout"""
        try:
            chart = self._find_chart(chart_name)
            if not chart:
                raise ValueError(f"Chart '{chart_name}' not found")
            
            chart.ApplyLayout(layout_id)
            
            return {"status": "success", "chart_name": chart_name, "layout_id": layout_id}
            
        except Exception as e:
            logger.error(f"Error applying chart layout: {e}")
            return {"status": "error", "message": str(e)}
    
    def set_chart_style(self, chart_name: str, style_id: int) -> Dict[str, Any]:
        """Apply predefined chart style"""
        try:
            chart = self._find_chart(chart_name)
            if not chart:
                raise ValueError(f"Chart '{chart_name}' not found")
            
            chart.ChartStyle = style_id
            
            return {"status": "success", "chart_name": chart_name, "style_id": style_id}
            
        except Exception as e:
            logger.error(f"Error setting chart style: {e}")
            return {"status": "error", "message": str(e)}
    
    def change_chart_type(self, chart_name: str, new_chart_type: str) -> Dict[str, Any]:
        """Change chart type at runtime"""
        try:
            chart = self._find_chart(chart_name)
            if not chart:
                raise ValueError(f"Chart '{chart_name}' not found")
            
            chart_type_value = getattr(ChartType, new_chart_type.upper()).value
            chart.ChartType = chart_type_value
            
            return {"status": "success", "chart_name": chart_name, "new_type": new_chart_type}
            
        except Exception as e:
            logger.error(f"Error changing chart type: {e}")
            return {"status": "error", "message": str(e)}
    
    def modify_pivot_fields(self, pivot_table_name: str, field_name: str, 
                           orientation: str, summary_function: Optional[str] = None) -> Dict[str, Any]:
        """Modify pivot table fields (add/remove/change)"""
        try:
            pivot_table = None
            for sheet in self.wb.sheets:
                try:
                    pivot_table = sheet.api.PivotTables(pivot_table_name)
                    break
                except:
                    continue
            
            if not pivot_table:
                raise ValueError(f"Pivot table '{pivot_table_name}' not found")
            
            field = pivot_table.PivotFields(field_name)
            
            # Set orientation
            orientation_map = {
                'ROW': 1, 'COLUMN': 2, 'PAGE': 3, 'DATA': 4, 'HIDDEN': 0
            }
            field.Orientation = orientation_map.get(orientation.upper(), 0)
            
            # Set summary function if it's a data field
            if orientation.upper() == 'DATA' and summary_function:
                function_map = {
                    'SUM': -4157, 'COUNT': -4112, 'AVERAGE': -4106,
                    'MAX': -4136, 'MIN': -4139, 'PRODUCT': -4149
                }
                field.Function = function_map.get(summary_function.upper(), -4157)
            
            return {"status": "success", "pivot_table": pivot_table_name, "field": field_name}
            
        except Exception as e:
            logger.error(f"Error modifying pivot fields: {e}")
            return {"status": "error", "message": str(e)}
    
    def create_calculated_field(self, pivot_table_name: str, field_name: str, formula: str) -> Dict[str, Any]:
        """Create calculated field in pivot table"""
        try:
            pivot_table = None
            for sheet in self.wb.sheets:
                try:
                    pivot_table = sheet.api.PivotTables(pivot_table_name)
                    break
                except:
                    continue
            
            if not pivot_table:
                raise ValueError(f"Pivot table '{pivot_table_name}' not found")
            
            calculated_field = pivot_table.CalculatedFields().Add(field_name, formula)
            
            return {"status": "success", "pivot_table": pivot_table_name, "calculated_field": field_name}
            
        except Exception as e:
            logger.error(f"Error creating calculated field: {e}")
            return {"status": "error", "message": str(e)}
    
    def add_slicer(self, pivot_table_name: str, field_name: str, position: Tuple[int, int] = (500, 100)) -> Dict[str, Any]:
        """Add slicer for pivot table filtering"""
        try:
            pivot_table = None
            target_sheet = None
            for sheet in self.wb.sheets:
                try:
                    pivot_table = sheet.api.PivotTables(pivot_table_name)
                    target_sheet = sheet
                    break
                except:
                    continue
            
            if not pivot_table:
                raise ValueError(f"Pivot table '{pivot_table_name}' not found")
            
            # Create slicer cache
            slicer_cache = self.wb.api.SlicerCaches.Add(pivot_table, field_name)
            
            # Add slicer to worksheet
            slicer = slicer_cache.Slicers.Add(target_sheet.api, "", field_name, position[0], position[1], 150, 200)
            
            return {"status": "success", "pivot_table": pivot_table_name, "slicer_field": field_name}
            
        except Exception as e:
            logger.error(f"Error adding slicer: {e}")
            return {"status": "error", "message": str(e)}
    
    def export_chart(self, chart_name: str, file_path: str, file_format: str = "PNG") -> Dict[str, Any]:
        """Export chart to file"""
        try:
            chart = self._find_chart(chart_name)
            if not chart:
                raise ValueError(f"Chart '{chart_name}' not found")
            
            # Export chart
            format_map = {"PNG": "PNG", "JPG": "JPG", "JPEG": "JPG", "GIF": "GIF", "PDF": "PDF"}
            export_format = format_map.get(file_format.upper(), "PNG")
            
            chart.Export(file_path, export_format)
            
            return {"status": "success", "chart_name": chart_name, "exported_to": file_path}
            
        except Exception as e:
            logger.error(f"Error exporting chart: {e}")
            return {"status": "error", "message": str(e)}
    
    def create_combo_chart(self, chart_name: str, data_range: str, 
                          primary_series: List[str], secondary_series: List[str],
                          primary_type: str = "COLUMN", secondary_type: str = "LINE") -> Dict[str, Any]:
        """Create combination chart with different chart types"""
        try:
            if not self.wb:
                raise ValueError("No workbook opened")
            
            # Create base chart
            source_range = self.ws.range(data_range)
            chart_obj = self.ws.api.ChartObjects().Add(100, 100, 500, 350)
            chart = chart_obj.Chart
            
            # Set data source
            chart.SetSourceData(source_range.api)
            
            # Configure primary series
            primary_type_value = getattr(ChartType, primary_type.upper()).value
            for i, series_name in enumerate(primary_series, 1):
                try:
                    series = chart.SeriesCollection(i)
                    series.ChartType = primary_type_value
                    series.AxisGroup = 1  # Primary axis
                except:
                    pass
            
            # Configure secondary series
            secondary_type_value = getattr(ChartType, secondary_type.upper()).value
            for i, series_name in enumerate(secondary_series, len(primary_series) + 1):
                try:
                    series = chart.SeriesCollection(i)
                    series.ChartType = secondary_type_value
                    series.AxisGroup = 2  # Secondary axis
                except:
                    pass
            
            return {
                "status": "success",
                "chart_name": chart.Name,
                "primary_type": primary_type,
                "secondary_type": secondary_type
            }
            
        except Exception as e:
            logger.error(f"Error creating combo chart: {e}")
            return {"status": "error", "message": str(e)}
    
    def get_chart_info(self, chart_name: str) -> Dict[str, Any]:
        """Get comprehensive chart information"""
        try:
            chart = self._find_chart(chart_name)
            if not chart:
                raise ValueError(f"Chart '{chart_name}' not found")
            
            info = {
                "status": "success",
                "chart_name": chart.Name,
                "chart_type": chart.ChartType,
                "has_title": chart.HasTitle,
                "title": chart.ChartTitle.Text if chart.HasTitle else None,
                "has_legend": chart.HasLegend,
                "legend_position": chart.Legend.Position if chart.HasLegend else None,
                "series_count": chart.SeriesCollection().Count,
                "data_source": str(chart.SeriesCollection(1).Formula) if chart.SeriesCollection().Count > 0 else None
            }
            
            return info
            
        except Exception as e:
            logger.error(f"Error getting chart info: {e}")
            return {"status": "error", "message": str(e)}
    
    def list_all_charts(self) -> Dict[str, Any]:
        """List all charts in the workbook"""
        try:
            charts = []
            for sheet in self.wb.sheets:
                for chart_obj in sheet.api.ChartObjects():
                    charts.append({
                        "name": chart_obj.Name,
                        "sheet": sheet.name,
                        "chart_type": chart_obj.Chart.ChartType
                    })
            
            return {"status": "success", "charts": charts}
            
        except Exception as e:
            logger.error(f"Error listing charts: {e}")
            return {"status": "error", "message": str(e)}
    
    def list_pivot_tables(self) -> Dict[str, Any]:
        """List all pivot tables in the workbook"""
        try:
            pivot_tables = []
            for sheet in self.wb.sheets:
                try:
                    for pt in sheet.api.PivotTables():
                        pivot_tables.append({
                            "name": pt.Name,
                            "sheet": sheet.name,
                            "source_data": str(pt.SourceData)
                        })
                except:
                    pass
            
            return {"status": "success", "pivot_tables": pivot_tables}
            
        except Exception as e:
            logger.error(f"Error listing pivot tables: {e}")
            return {"status": "error", "message": str(e)}
    
    def _find_chart(self, chart_name: str):
        """Helper method to find chart by name across all sheets"""
        for sheet in self.wb.sheets:
            try:
                for chart_obj in sheet.api.ChartObjects():
                    if chart_obj.Name == chart_name:
                        return chart_obj.Chart
            except:
                continue
        return None
    
    def close_workbook(self, save: bool = True) -> Dict[str, Any]:
        """Close workbook and cleanup"""
        try:
            if self.wb:
                if save:
                    self.wb.save()
                self.wb.close()
                self.wb = None
                self.ws = None
            
            if self.app:
                self.app.quit()
                self.app = None
            
            return {"status": "success", "message": "Workbook closed successfully"}
            
        except Exception as e:
            logger.error(f"Error closing workbook: {e}")
            return {"status": "error", "message": str(e)}

# Utility functions for common chart operations
def create_standard_pivot_chart(workbook_path: str, data_range: str, 
                               row_fields: List[str], value_fields: List[str],
                               chart_type: str = "COLUMN", chart_title: str = None) -> Dict[str, Any]:
    """One-stop function to create a standard pivot chart"""
    core = ExcelChartsCore(workbook_path)
    try:
        # Create pivot table
        pt_result = core.create_pivot_table(
            data_range=data_range,
            pivot_table_name="AutoPivotTable",
            destination="H1",
            row_fields=row_fields,
            value_fields=value_fields
        )
        
        if pt_result["status"] != "success":
            return pt_result
        
        # Create chart from pivot table
        chart_result = core.create_pivot_chart_from_table(
            pivot_table_name="AutoPivotTable",
            chart_type=chart_type,
            chart_title=chart_title
        )
        
        return chart_result
        
    finally:
        core.close_workbook()

def create_dashboard_charts(workbook_path: str, chart_configs: List[Dict]) -> Dict[str, Any]:
    """Create multiple charts for dashboard"""
    core = ExcelChartsCore(workbook_path)
    results = []
    
    try:
        for config in chart_configs:
            if config.get("type") == "pivot":
                result = core.create_pivot_chart_from_table(**config["params"])
            else:
                result = core.create_chart_from_range(**config["params"])
            results.append(result)
        
        return {"status": "success", "charts_created": len([r for r in results if r["status"] == "success"])}
        
    finally:
        core.close_workbook()