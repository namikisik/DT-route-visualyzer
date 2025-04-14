import pandas as pd
import folium
import os
import shutil
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
# from openpyxl import load_workbook
# from openpyxl.drawing.image import Image
import xlwings as xw


# Constants
TEMPLATE_PATH = 'svt_temp.xlsx'
OUTPUT_PATH = 'svt_report.xlsx'
DATA_PATH = 'Export.csv'
SCREENSHOT_PATH = 'map_screenshot.png'
HTML_PATH = 'signal_map.html'

# 1. Data Loading and Preparation
def load_and_prepare_data():
    try:
        data = pd.read_csv(DATA_PATH)
        data = data.set_index("Time").groupby(level=0).first()
        # data = data[data['Band (LTE pcell)']=='LTE FDD 2100 band 1']
        df = pd.DataFrame({
            'rsrp': data["RSRP (LTE pcell)"],
            'lat': data["Latitude"],
            'lon': data["Longitude"]
        }).dropna(subset=['rsrp'])
        
        return df
    except Exception as e:
        print(f"Error loading data: {str(e)}")
        exit()

# 2. RSRP Configuration
RSRP_CONFIG = {
    'Excellent (≥ -80)': ('#006400', -80, 100),
    'Good (-90 to -80)': ('#008000', -90, -80),
    'Fair (-100 to -90)': ('#FFA500', -100, -90),
    'Poor (-110 to -100)': ('#FF0000', -110, -100),
    'Bad (< -110)': ('#8B0000', -200, -110)
}

def categorize_rsrp_data(df):
    def categorize_rsrp(rsrp):
        for label, (color, min_val, max_val) in RSRP_CONFIG.items():
            if min_val <= rsrp < max_val:
                return (label, color)
        return ('Bad (< -110)', '#8B0000')

    df['category'], df['color'] = zip(*df['rsrp'].apply(categorize_rsrp))
    return df, df['category'].value_counts(normalize=True).mul(100).round(1)

# 3. Map Creation
def create_map(df, center_location):
    try:
        m = folium.Map(
            location=center_location,
            zoom_start=14,
            tiles='CartoDB positron',
            attr='',
            control_scale=False
        )

        # Draw route with thick lines
        for i in range(len(df)-1):
            folium.PolyLine(
                locations=df[['lat', 'lon']].iloc[i:i+2],
                color=df['color'].iloc[i],
                weight=8,
                opacity=0.9
            ).add_to(m)

        return m
    except Exception as e:
        print(f"Error creating map: {str(e)}")
        exit()

def add_legend(map_obj, percentages):
    legend_html = f'''
    <div style="position: fixed; bottom: 25px; left: 25px; z-index: 1000;
                background: rgba(255,255,255,0.95); padding: 18px;
                border-radius: 10px; box-shadow: 0 2px 8px rgba(0,0,0,0.15);
                font-family: Helvetica, Arial; width: 260px;">
        <h4 style="margin: 0 0 15px 0; color: #2d3436; font-size: 18px;
                   font-weight: 500;">
            Signal Quality Distribution
        </h4>
        <div style="padding-left: 10px;">
            {'<br>'.join([
                f'<div style="margin: 12px 0; display: flex; align-items: center">'
                f'<div style="width: 20px; height: 20px; background: {RSRP_CONFIG[label][0]};'
                f'margin-right: 12px; border: 1px solid #ddd;"></div>'
                f'<div>'
                f'<div style="color: #2d3436; font-size: 15px; margin-bottom: 4px;">'
                f'{label}</div>'
                f'<div style="color: #636e72; font-size: 14px;">'
                f'- {percentages.get(label, 0.0):.1f}%</div>'
                f'</div></div>'
                for label in RSRP_CONFIG.keys()
            ])}
        </div>
    </div>
    '''
    map_obj.get_root().html.add_child(folium.Element(legend_html))
    return map_obj

# 4. Screenshot Capture
def capture_screenshot(html_file):
    try:
        options = webdriver.ChromeOptions()
        options.add_argument('--headless')
        options.add_argument('--disable-gpu')
        options.add_argument('--window-size=1200,800')
        options.add_argument('--force-device-scale-factor=1.5')
        options.add_argument('--hide-scrollbars')

        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)

        driver.get(f'file:///{html_file}')
        
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CLASS_NAME, 'folium-map')))
        
        WebDriverWait(driver, 20).until(
            lambda d: d.execute_script('return document.readyState === "complete"'))
        
        # Remove any UI elements
        driver.execute_script("""
            const elements = document.querySelectorAll(
                '.leaflet-control-container, .folium-map'
            );
            elements.forEach(el => el.style.display = 'block');
            document.body.style.overflow = 'hidden';
        """)
        
        map_element = driver.find_element(By.CLASS_NAME, 'folium-map')
        map_element.screenshot(SCREENSHOT_PATH)
        print(f"Screenshot saved to {SCREENSHOT_PATH}")
        return True

    except Exception as e:
        print(f"Error capturing screenshot: {str(e)}")
        return False
    finally:
        if 'driver' in locals():
            driver.quit()

# 5. Excel Integration
def insert_into_excel():
    try:
        # Copy template to output path
        shutil.copyfile(TEMPLATE_PATH, OUTPUT_PATH)

        # Load workbook and select sheet
        wb = xw.Book(OUTPUT_PATH)
        ws = wb.sheets['deneme']

        # Define target cell range
        cell_range = ws.range('C5:F13')

        # Calculate dimensions in points
        left = cell_range.left
        top = cell_range.top
        width = cell_range.width
        height = cell_range.height

        # Insert image
        ws.pictures.add(SCREENSHOT_PATH, 
                        left=left, 
                        top=top, 
                        width=width, 
                        height=height)

        # Save and close
        wb.save(OUTPUT_PATH)
        wb.close()

        print(f"Report successfully saved to {OUTPUT_PATH}")
        return True

    except Exception as e:
        print(f"Error inserting into Excel: {str(e)}")
        return False


# Main execution
def main():
    # 1. Load and prepare data
    df = load_and_prepare_data()
    
    # 2. Categorize RSRP data
    df, percentages = categorize_rsrp_data(df)
    
    # 3. Create and save map
    center_location = [df['lat'].mean(), df['lon'].mean()]
    m = create_map(df, center_location)
    m = add_legend(m, percentages)
    m.save(HTML_PATH)
    print(f"Map saved to {HTML_PATH}")
    
    # 4. Capture screenshot
    if not capture_screenshot(os.path.abspath(HTML_PATH)):
        exit()
    
    # 5. Insert into Excel
    if not insert_into_excel():
        exit()
    
    print("Process completed successfully!")

if __name__ == "__main__":
    main()
