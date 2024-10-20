import requests
import pandas as pd
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from tqdm import tqdm
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# URL для PageSpeed Insights API
API_URL = "https://www.googleapis.com/pagespeedonline/v5/runPagespeed"

# Ключ API (его нужно получить в Google Cloud Console)
API_KEY = "YOUR_API_KEY"

def read_sites(file_path):
    with open(file_path, "r") as file:
        sites = [line.strip() for line in file if line.strip()]
    return sites

# Функция для запроса PageSpeed данных
def fetch_lighthouse_data(url, strategy):
    params = {
        'url': url,
        'strategy': strategy,
        'key': API_KEY
    }
    response = requests.get(API_URL, params=params)
    if response.status_code == 200:
        return response.json()
    else:
        return None

# Извлечение нужных метрик из JSON
def extract_metrics(data):
    try:
        metrics = {
            'Score': int(data['lighthouseResult']['categories']['performance']['score'] * 100),
            'FCP': round(data['lighthouseResult']['audits']['first-contentful-paint']['numericValue'] / 1000, 1),
            'LCP': round(data['lighthouseResult']['audits']['largest-contentful-paint']['numericValue'] / 1000, 1),
            'SI': round(data['lighthouseResult']['audits']['speed-index']['numericValue'] / 1000, 1),
            'TBT': int(data['lighthouseResult']['audits']['total-blocking-time']['numericValue']),
            'CLS': round(data['lighthouseResult']['audits']['cumulative-layout-shift']['numericValue'], 3),
            'TTFB': round(data['lighthouseResult']['audits']['server-response-time']['numericValue'] / 1000, 1),
            'INP': data['lighthouseResult']['audits'].get('interaction-to-next-paint', {}).get('numericValue', None)
        }
        return metrics
    except (KeyError, TypeError):
        return None

def check_site(site, device, attempts=10):
    results = []
    for i in range(attempts):
        data = fetch_lighthouse_data(site, device)
        if data:
            metrics = extract_metrics(data)
            if metrics:
                results.append(metrics)
        time.sleep(1)
    return results

def calculate_average(results):
    averages = {}
    metrics = ['Score', 'FCP', 'LCP', 'SI', 'TBT', 'CLS', 'TTFB', 'INP']
    for metric in metrics:
        valid_values = [result[metric] for result in results if result.get(metric) is not None]
        if metric in ['FCP', 'LCP', 'SI', 'TTFB']:
            averages[metric] = round(sum(valid_values) / len(valid_values), 1) if valid_values else None
        elif metric == 'TBT':
            averages[metric] = int(sum(valid_values) / len(valid_values)) if valid_values else None
        elif metric == 'CLS':
            averages[metric] = round(sum(valid_values) / len(valid_values), 3) if valid_values else None
        else:
            averages[metric] = int(sum(valid_values) / len(valid_values)) if valid_values and metric == 'Score' else (sum(valid_values) / len(valid_values) if valid_values else None)
    return averages

def collect_data_for_sites(sites):
    all_results = []
    summary_results = []

    with ThreadPoolExecutor(max_workers=10) as executor:
        future_to_site = {}
        tasks = []

        for site in sites:
            for device in ['desktop', 'mobile']:
                future = executor.submit(check_site, site, device)
                future_to_site[future] = (site, device)
                tasks.append(future)

        for future in tqdm(as_completed(tasks), total=len(tasks), desc="Обработка URL"):
            site, device = future_to_site[future]
            try:
                device_results = future.result()
                
                if device_results:
                    average_metrics = calculate_average(device_results)

                    for result in device_results:
                        all_results.append({
                            'site': site,
                            'device': device,
                            **result
                        })
                    
                    summary_results.append({
                        'site': site,
                        'device': device,
                        **average_metrics
                    })
            except Exception as e:
                print(f"Ошибка при обработке {site} на {device}: {e}")

    return all_results, summary_results

def save_to_excel(sites_data, summary_data, file_name='results.xlsx'):
    checks_df = pd.DataFrame(sites_data)
    summary_df = pd.DataFrame(summary_data)

    with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
        checks_df.to_excel(writer, sheet_name='Проверка страниц', index=False)
        summary_df.to_excel(writer, sheet_name='Среднее значение', index=False)

    # Открываем файл для форматирования с помощью openpyxl
    wb = load_workbook(file_name)
    ws = wb['Среднее значение']

    # Добавляем единицы измерения в шапку
    ws['C1'] = 'Score'
    ws['D1'] = 'FCP, сек.'
    ws['E1'] = 'LCP, сек.'
    ws['F1'] = 'SI, сек.'
    ws['G1'] = 'TBT, мс.'
    ws['H1'] = 'CLS'
    ws['I1'] = 'TTFB, сек.'
    ws['J1'] = 'INP, мс.'

    # Условное форматирование для заливки цветов
    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")  # Допустимые значения
    orange_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")  # Средние значения
    red_fill = PatternFill(start_color="F2DCDB", end_color="F2DCDB", fill_type="solid")  # Критические значения

    # Определение диапазона значений
    for row in range(2, ws.max_row + 1):
        # Score
        score_value = ws[f'C{row}'].value
        if score_value is not None:
            if score_value >= 90:
                ws[f'C{row}'].fill = green_fill
            elif 50 <= score_value < 90:
                ws[f'C{row}'].fill = orange_fill
            else:
                ws[f'C{row}'].fill = red_fill

        # FCP
        fcp_value = ws[f'D{row}'].value
        if fcp_value is not None:
            if fcp_value <= 1.8:
                ws[f'D{row}'].fill = green_fill
            elif 1.8 < fcp_value <= 3.0:
                ws[f'D{row}'].fill = orange_fill
            else:
                ws[f'D{row}'].fill = red_fill
        
        # LCP
        lcp_value = ws[f'E{row}'].value
        if lcp_value is not None:
            if lcp_value <= 2.5:
                ws[f'E{row}'].fill = green_fill
            elif 2.5 < lcp_value <= 4.0:
                ws[f'E{row}'].fill = orange_fill
            else:
                ws[f'E{row}'].fill = red_fill
        
        # SI
        si_value = ws[f'F{row}'].value
        if si_value is not None:
            if si_value <= 3.4:
                ws[f'F{row}'].fill = green_fill
            elif 3.4 < si_value <= 5.8:
                ws[f'F{row}'].fill = orange_fill
            else:
                ws[f'F{row}'].fill = red_fill
        
        # TBT
        tbt_value = ws[f'G{row}'].value
        if tbt_value is not None:
            if tbt_value <= 200:
                ws[f'G{row}'].fill = green_fill
            elif 200 < tbt_value <= 600:
                ws[f'G{row}'].fill = orange_fill
            else:
                ws[f'G{row}'].fill = red_fill
        
        # CLS
        cls_value = ws[f'H{row}'].value
        if cls_value is not None:
            if cls_value <= 0.1:
                ws[f'H{row}'].fill = green_fill
            elif 0.1 < cls_value <= 0.25:
                ws[f'H{row}'].fill = orange_fill
            else:
                ws[f'H{row}'].fill = red_fill
        
        # TTFB
        ttfb_value = ws[f'I{row}'].value
        if ttfb_value is not None:
            if ttfb_value <= 0.8:
                ws[f'I{row}'].fill = green_fill
            elif 0.8 < ttfb_value <= 1.8:
                ws[f'I{row}'].fill = orange_fill
            else:
                ws[f'I{row}'].fill = red_fill
        
        # INP
        inp_value = ws[f'J{row}'].value
        if inp_value is not None:
            if inp_value <= 200:
                ws[f'J{row}'].fill = green_fill
            elif 200 < inp_value <= 500:
                ws[f'J{row}'].fill = orange_fill
            else:
                ws[f'J{row}'].fill = red_fill

    # Сохраняем изменения
    wb.save(file_name)
    wb.close()

def main():
    sites = read_sites('site.txt')
    
    all_results, summary_results = collect_data_for_sites(sites)

    save_to_excel(all_results, summary_results)

if __name__ == "__main__":
    main()