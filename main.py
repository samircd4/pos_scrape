from playwright.sync_api import sync_playwright
from rich import print
from rich.console import Console
from rich.table import Table
import pandas as pd

console = Console()

def save_data(data):
    try:
        old_df = pd.read_excel('orders.xlsx')
    except FileNotFoundError:
        old_df = pd.DataFrame()
    
    df = pd.DataFrame(data)
    new_df = pd.concat([old_df, df])
    new_df.to_excel('orders.xlsx', index=False)
    return

def read_data():
    try:
        df = pd.read_excel('orders.xlsx')
    except FileNotFoundError:
        return []
    
    oerders = []
    for index, row in df.iterrows():
        order_id = row['order_id']
        if str(order_id) in oerders:
            continue
        oerders.append(row['order_id'])
    return oerders

def get_orders(day, month, year, order_type):
    url = 'https://wdfpos.com/reports'
    console.log('Navigated to', url)
    orders = []
    existing_orders = read_data()
    with sync_playwright() as p:
        browser = p.firefox.launch(headless=False, slow_mo=2000)
        context = browser.new_context(storage_state='login.json')
        context.set_default_timeout(60000)
        page = context.new_page()
        console.log('Logining in...')
        page.goto(url)
        # page.wait_for_timeout(15000)
        # email_input = page.locator('input.baTaHaLaE').fill('maxsabo@noltra.com')
        # pass_input = page.locator('input.baTaHaLaV').fill('SomedayVerySoon408$')
        
        # login_button = page.locator('div.fa-lock').click()
        page.wait_for_timeout(2000)
        console.log('Logged in successfully')
        page.get_by_role("combobox").first.select_option('\"Order List\"')
        
        # page.select_option('/html/body/div[1]/div[2]/div/div[2]/div[3]/div[2]/select[5]', value='Laundry Orders Only')
        if '1' in order_type:
            page.get_by_role("combobox").last.select_option('\"Laundry Orders Only\"')
            order_type = 'Laundry'
        else:
            page.get_by_role("combobox").last.select_option('\"Retail Orders Only\"')
            order_type = 'Retail'
        
        page.query_selector_all('div.date_picker_element')[0].click()
        page.select_option("select.picker__select--year", value=year)
        page.select_option("select.picker__select--month", value=month)
        page.click(f'div.picker__day.picker__day--infocus:has-text("{day}")')
        console.log('Start date selected')
        input('Select end date and press enter to continue...')
        
        # page.query_selector_all('div.date_picker_element')[1].click()
        # page.query_selector_all('div.picker__nav--prev')[1].click()
        # page.get_by_role("row", name="Sun, 9/29/24 Mon, 9/30/24 Tue").get_by_label("Mon, 9/30/").click()
        # console.log('End date selected')
        
        page.get_by_role("button", name="Display Report").click()
        page.wait_for_timeout(2000)
        console.log('Report displayed')
        
        total_orders = int(page.locator('div.cmaSlaF').inner_text().split(' ')[0])
        display_orders = page.locator('div.cmaSlaF').inner_text().split(')')[0].split(' ')[-1]

        while total_orders > int(display_orders):
            page.locator('div.cmaSlaF').click()
            console.log(f'Loading...{display_orders}/{total_orders}')
            page.wait_for_timeout(1000)
            total_orders = int(page.locator('div.cmaSlaF').inner_text().split(' ')[0])
            display_orders = page.locator('div.cmaSlaF').inner_text().split(')')[0].split(' ')[-1]
            # if int(display_orders) >= int(60):
        
        items = page.query_selector_all('div.cmaThaD > div')
        current_num = 0
        for i in items:
            table = Table(show_header=True, header_style="bold magenta")
            table.add_column("created_date", justify="right", style="cyan", no_wrap=True)
            table.add_column("Order ID", style="dim", justify="right")
            table.add_column("Order Type", style="dim", justify="right")
            table.add_column("customer", style="magenta")
            table.add_column("quantity_details", style="magenta")
            table.add_column("unit_price", style="magenta")
            table.add_column("pay_status", justify="right", style="green")
            table.add_column("subtotal", justify="right", style="green")
            try:
                order = {}
                current_num += 1
                created_date = i.query_selector('div.cmaThaG').inner_text()
                order_id = i.query_selector('div.cmywaJ0').inner_text()
                
                if int(order_id) in existing_orders:
                    console.log(f'{order_id} Order id is already scraped')
                    continue
                
                customer = i.query_selector('div.cmaThaJ').inner_text()
                company = i.query_selector('div.cmaThaK').inner_text()
                order_status = i.query_selector('div.cmaThaI').inner_text()
                pay_status = i.query_selector('div.cmaThaL').inner_text()
                pay_method = i.query_selector('div.cmaThaM').inner_text()
                cc_tip = i.query_selector('div.cmaThaP').inner_text()
                
                
                order['created_date'] = created_date
                order['order_id'] = order_id
                order['order_type'] = order_type
                order['customer'] = customer
                order['company'] = company
                order['order_status'] = order_status
                order['pay_status'] = pay_status
                order['pay_method'] = pay_method
                order['cc_tip'] = cc_tip
                
                
                button = i.query_selector('button.cmaYaNaX1')
            except Exception as e:
                print(f'{order_id} Order id is not scraped due to {e}')
                save_data(orders)
                return
            try:
                if button:
                    button.click()
                    page.wait_for_timeout(1000)
                    products_len = i.query_selector_all('div#line-item-group')
                    
                    for p in products_len:
                        order_copy = order.copy()
                        qty = p.query_selector('div.cmaYaQaZ0').inner_text()
                        p_name = p.query_selector('div.cmaYaQc0').inner_text()
                        quantity_details = f'{qty.strip()} {p_name.strip()}'
                        note = p.query_selector('div.cmaYaQf0').inner_text()
                        discount = p.query_selector('div.cmaYaVm').inner_text()
                        unit_price = p.query_selector('div.cmaYaQi0').inner_text()
                        tax = p.query_selector('div.cmaYaVj').inner_text()
                        subtotal = p.query_selector('div.cmaYaQl0').inner_text()
                        
                        
                        order_copy['quantity_details'] = quantity_details
                        order_copy['note'] = note
                        order_copy['discount'] = discount
                        order_copy['unit_price'] = unit_price
                        order_copy['tax'] = tax
                        order_copy['subtotal'] = subtotal
                        orders.append(order_copy)
                        table.add_row(created_date,order_id,order_type, customer,quantity_details, unit_price, pay_status, subtotal)
                        # print(order)
            except Exception as e:
                print(f'{order_id} Order id is not scraped due to {e}')
                with open('error.txt', 'a') as f:
                    f.write(f'{order_id} Order id is not scraped due to {e}\n')
                save_data(orders)
                continue
            
            console.print(table)
            console.log(f'{current_num}/{len(items)} -> Date: {int(month)+1}-{day}-{year} <-')
        page.wait_for_timeout(1000)
    
    save_data(orders)
    return orders


if __name__ == '__main__':
    print("""
Welcome to the Order Scraper
for date: please do not enter the zero in front of the date.
The scraper will scrape order from the given date to todays date.
For order type:
Enter 1 for Laundry
Enter 2 for Retail
""")
    date_str = input('Enter start date (MM-DD-YYYY): ')
    order_type = input('Enter order type: ')
    
    month, day, year = date_str.split('-')
    month = str(int(month)-1)
    old_orders = read_data()
    console.log(f'{len(old_orders)} Items are already scraped')
    unique_orders = []
    for o in old_orders:
        if str(o) not in unique_orders:
            unique_orders.append(str(o))
    console.log(f'{len(unique_orders)} Orders are unique')
    get_orders(day, month, year, order_type)
    console.log('Scraping Done!')