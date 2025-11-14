"""
ExtJS Helper Functions
Provides utilities to interact with ExtJS components using JavaScript API only
No DOM clicking, no <li> searching, pure ExtJS ComponentQuery
"""

import time
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException


class ExtJSHelper:
    """Helper class for ExtJS component interactions using JavaScript API"""
    
    def __init__(self, driver: WebDriver, wait: WebDriverWait = None):
        """
        Initialize ExtJS helper
        
        Args:
            driver: Selenium WebDriver instance
            wait: Optional WebDriverWait instance
        """
        self.driver = driver
        self.wait = wait or WebDriverWait(driver, 15)
    
    def check_extjs_available(self) -> bool:
        """
        Check if ExtJS is available in the current context
        
        Returns:
            True if ExtJS is available
        """
        try:
            js_code = """
            (function() {
                try {
                    return typeof Ext !== 'undefined' && typeof Ext.ComponentQuery !== 'undefined';
                } catch (e) {
                    return false;
                }
            })();
            """
            result = self.driver.execute_script(js_code)
            return result is True
        except Exception as e:
            # JavaScript execution might fail if iframe is not ready
            return False
    
    def get_extjs_combo_values(self, component_name: str) -> list:
        """
        Return all values inside an ExtJS combobox store
        
        Args:
            component_name: Name of the ExtJS combobox component (e.g., "ctl00$ContentPlaceHolder1$ddlBank")
            
        Returns:
            List of string values available in the combobox
        """
        js_code = f"""
        (function() {{
            try {{
                if (typeof Ext === 'undefined') {{
                    return {{success: false, error: 'ExtJS not available'}};
                }}
                
                // Find combobox by name
                var combo = Ext.ComponentQuery.query("combobox[name='{component_name}']")[0];
                
                if (!combo) {{
                    // Try alternative: search all comboboxes for matching name
                    var allCombos = Ext.ComponentQuery.query("combobox");
                    for (var i = 0; i < allCombos.length; i++) {{
                        if (allCombos[i].name === '{component_name}') {{
                            combo = allCombos[i];
                            break;
                        }}
                    }}
                }}
                
                if (!combo) {{
                    return {{success: false, error: 'Combobox not found: {component_name}'}};
                }}
                
                // Get store and extract all values
                var store = combo.getStore();
                if (!store) {{
                    return {{success: false, error: 'Store not available'}};
                }}
                
                var values = [];
                var displayField = combo.displayField || 'text';
                var valueField = combo.valueField || 'value';
                
                store.each(function(record) {{
                    var displayValue = record.get(displayField);
                    var actualValue = record.get(valueField);
                    // Use display value if available, otherwise use value
                    values.push(displayValue || actualValue || '');
                }});
                
                return {{
                    success: true,
                    values: values,
                    count: values.length
                }};
            }} catch (e) {{
                return {{
                    success: false,
                    error: 'Exception: ' + e.toString() + ' - ' + e.message
                }};
            }}
        }})();
        """
        
        result = self.driver.execute_script(js_code)
        
        if result and result.get('success'):
            return result.get('values', [])
        else:
            error = result.get('error', 'Unknown error') if result else 'No result returned'
            print(f"[WARNING] Failed to get combo values: {error}")
            return []
    
    def set_extjs_combo(self, component_name: str, value: str) -> bool:
        """
        Set ExtJS combobox value using setValue() and fireEvent('select')
        No DOM clicking, pure ExtJS API
        
        Args:
            component_name: Name of the ExtJS combobox component
            value: Value to set (must match one of the values in the store)
            
        Returns:
            True if successful
        """
        js_code = f"""
        (function() {{
            try {{
                if (typeof Ext === 'undefined') {{
                    return {{success: false, error: 'ExtJS not available'}};
                }}
                
                // Find combobox by name
                var combo = Ext.ComponentQuery.query("combobox[name='{component_name}']")[0];
                
                if (!combo) {{
                    // Try alternative: search all comboboxes
                    var allCombos = Ext.ComponentQuery.query("combobox");
                    for (var i = 0; i < allCombos.length; i++) {{
                        if (allCombos[i].name === '{component_name}') {{
                            combo = allCombos[i];
                            break;
                        }}
                    }}
                }}
                
                if (!combo) {{
                    return {{success: false, error: 'Combobox not found: {component_name}'}};
                }}
                
                // Set the value
                combo.setValue('{value}');
                
                // Fire select event to trigger PostBack
                combo.fireEvent('select', combo, combo.getStore().findRecord(combo.valueField || 'value', '{value}'));
                
                // Also fire change event for good measure
                combo.fireEvent('change', combo, '{value}');
                
                return {{
                    success: true,
                    comboId: combo.id || 'unknown',
                    setValue: '{value}'
                }};
            }} catch (e) {{
                return {{
                    success: false,
                    error: 'Exception: ' + e.toString() + ' - ' + e.message
                }};
            }}
        }})();
        """
        
        result = self.driver.execute_script(js_code)
        
        if result and result.get('success'):
            print(f"[OK] Set combobox '{component_name}' to '{value}'")
            time.sleep(0.1)  # Reduced from 0.5s to 0.1s (80% faster) - Wait for PostBack
            return True
        else:
            error = result.get('error', 'Unknown error') if result else 'No result returned'
            print(f"[ERROR] Failed to set combo value: {error}")
            return False
    
    def click_tampilkan(self) -> bool:
        """
        Click "Tampilkan" button using ExtJS API
        Finds button via ComponentQuery and triggers click event
        
        Returns:
            True if successful
        """
        js_code = """
        (function() {
            try {
                if (typeof Ext === 'undefined') {
                    return {success: false, error: 'ExtJS not available'};
                }
                
                // Find button by text or value
                var buttons = Ext.ComponentQuery.query("button");
                var tampilkanBtn = null;
                
                for (var i = 0; i < buttons.length; i++) {
                    var btn = buttons[i];
                    var text = btn.text || btn.getText() || '';
                    var value = btn.value || '';
                    
                    if (text.toLowerCase().indexOf('tampilkan') !== -1 || 
                        value.toLowerCase().indexOf('tampilkan') !== -1) {
                        tampilkanBtn = btn;
                        break;
                    }
                }
                
                // Also try to find by component type
                if (!tampilkanBtn) {
                    var allComponents = Ext.ComponentQuery.query("*");
                    for (var i = 0; i < allComponents.length; i++) {
                        var comp = allComponents[i];
                        if (comp.xtype === 'button' || comp instanceof Ext.button.Button) {
                            var text = comp.text || comp.getText() || '';
                            if (text.toLowerCase().indexOf('tampilkan') !== -1) {
                                tampilkanBtn = comp;
                                break;
                            }
                        }
                    }
                }
                
                if (!tampilkanBtn) {
                    return {success: false, error: 'Tampilkan button not found'};
                }
                
                // Trigger click event
                tampilkanBtn.fireEvent('click', tampilkanBtn);
                
                // Also try handler if available
                if (tampilkanBtn.handler) {
                    tampilkanBtn.handler.call(tampilkanBtn.scope || tampilkanBtn, tampilkanBtn);
                }
                
                return {
                    success: true,
                    buttonId: tampilkanBtn.id || 'unknown'
                };
            } catch (e) {
                return {
                    success: false,
                    error: 'Exception: ' + e.toString() + ' - ' + e.message
                };
            }
        })();
        """
        
        result = self.driver.execute_script(js_code)
        
        if result and result.get('success'):
            print("[OK] Clicked 'Tampilkan' button using ExtJS API")
            time.sleep(0.2)  # Reduced from 1.0s to 0.2s (80% faster) - Wait for form submission
            return True
        else:
            error = result.get('error', 'Unknown error') if result else 'No result returned'
            print(f"[ERROR] Failed to click Tampilkan: {error}")
            return False
    
    def wait_for_grid(self, timeout: int = 30) -> bool:
        """
        Wait until ExtJS grid is loaded and available
        
        Args:
            timeout: Maximum time to wait in seconds
            
        Returns:
            True if grid is found
        """
        js_code = """
        (function() {
            if (typeof Ext === 'undefined') {
                return false;
            }
            
            var grids = Ext.ComponentQuery.query("grid");
            if (grids && grids.length > 0) {
                var grid = grids[0];
                // Check if grid has data
                var store = grid.getStore();
                if (store && store.isLoaded && store.isLoaded()) {
                    return true;
                }
            }
            return false;
        })();
        """
        
        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                if self.driver.execute_script(js_code):
                    print("[OK] ExtJS grid is loaded")
                    return True
            except:
                pass
            time.sleep(0.1)  # Reduced from 0.5s to 0.1s (80% faster)
        
        print(f"[WARNING] Grid not found after {timeout} seconds")
        return False
    
    def get_grid_data(self) -> list:
        """
        Extract data from ExtJS grid using store
        
        Returns:
            List of dictionaries representing grid rows
        """
        js_code = """
        (function() {
            try {
                if (typeof Ext === 'undefined') {
                    return {success: false, error: 'ExtJS not available'};
                }
                
                var grids = Ext.ComponentQuery.query("grid");
                if (!grids || grids.length === 0) {
                    return {success: false, error: 'No grid found'};
                }
                
                var grid = grids[0];
                var store = grid.getStore();
                if (!store) {
                    return {success: false, error: 'Grid store not available'};
                }
                
                var columns = grid.columns || [];
                var columnFields = [];
                for (var i = 0; i < columns.length; i++) {
                    var col = columns[i];
                    if (col.dataIndex) {
                        columnFields.push(col.dataIndex);
                    }
                }
                
                var data = [];
                store.each(function(record) {
                    var row = {};
                    for (var j = 0; j < columnFields.length; j++) {
                        var field = columnFields[j];
                        row[field] = record.get(field) || '';
                    }
                    // Also get all fields if columnFields is empty
                    if (columnFields.length === 0) {
                        var fields = record.getFields();
                        fields.each(function(field) {
                            row[field.name] = record.get(field.name) || '';
                        });
                    }
                    data.push(row);
                });
                
                return {
                    success: true,
                    data: data,
                    rowCount: data.length,
                    columnFields: columnFields
                };
            } catch (e) {
                return {
                    success: false,
                    error: 'Exception: ' + e.toString() + ' - ' + e.message
                };
            }
        })();
        """
        
        result = self.driver.execute_script(js_code)
        
        if result and result.get('success'):
            return result.get('data', [])
        else:
            error = result.get('error', 'Unknown error') if result else 'No result returned'
            print(f"[WARNING] Failed to get grid data: {error}")
            return []
    
    def find_combo_by_position(self, position: int) -> str:
        """
        Find combobox by position (0=first, 1=second, etc.)
        Useful when component names are not known
        
        Args:
            position: Position index (0-based)
            
        Returns:
            Component name if found, empty string otherwise
        """
        js_code = f"""
        (function() {{
            try {{
                if (typeof Ext === 'undefined') {{
                    return {{success: false, error: 'ExtJS not available'}};
                }}
                
                var combos = Ext.ComponentQuery.query("combobox");
                if (combos && combos.length > {position}) {{
                    var combo = combos[{position}];
                    return {{
                        success: true,
                        name: combo.name || '',
                        id: combo.id || '',
                        value: combo.getValue() || ''
                    }};
                }}
                return {{success: false, error: 'Combobox at position {position} not found'}};
            }} catch (e) {{
                return {{success: false, error: 'Exception: ' + e.toString()}};
            }}
        }})();
        """
        
        result = self.driver.execute_script(js_code)
        
        if result and result.get('success'):
            return result.get('name', '')
        return ''
    
    def list_all_combos(self) -> list:
        """
        List all available comboboxes with their names and IDs
        Useful for debugging
        
        Returns:
            List of dictionaries with combo info
        """
        js_code = """
        (function() {
            try {
                if (typeof Ext === 'undefined') {
                    return {success: false, error: 'ExtJS not available'};
                }
                
                var combos = Ext.ComponentQuery.query("combobox");
                var comboList = [];
                
                for (var i = 0; i < combos.length; i++) {
                    var combo = combos[i];
                    comboList.push({
                        index: i,
                        name: combo.name || '',
                        id: combo.id || '',
                        inputId: combo.inputId || '',
                        value: combo.getValue() || ''
                    });
                }
                
                return {
                    success: true,
                    combos: comboList,
                    count: comboList.length
                };
            } catch (e) {
                return {
                    success: false,
                    error: 'Exception: ' + e.toString()
                };
            }
        })();
        """
        
        result = self.driver.execute_script(js_code)
        
        if result and result.get('success'):
            return result.get('combos', [])
        return []

