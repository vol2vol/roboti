import os
import sys
import time
import re
from dataclasses import dataclass
from typing import List, Optional

from selenium import webdriver
from selenium.common.exceptions import (
    SessionNotCreatedException,
    TimeoutException,
    ElementNotInteractableException,
    ElementClickInterceptedException,
    NoSuchElementException,
)
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager


USERNAME = "problem_user"     
PASSWORD = "secret_sauce"
SORT_VALUE = "az"             
PICK_STRATEGY = "last"        
CART_ACTION = "checkout"      


@dataclass
class InventoryItem:
    root: object
    name: str
    price: float
    add_btn: object


def start_opera_driver():
    opera_binary = os.environ.get("OPERA_BINARY")
    if not opera_binary:
        raise RuntimeError("Не найден OPERA_BINARY. Укажите путь к opera.exe")

    def build_driver(driver_version: Optional[str] = None):
        options = ChromeOptions()
        options.binary_location = opera_binary
        options.add_argument("--start-maximized")
        options.add_argument("--disable-infobars")
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")
        service = ChromeService(ChromeDriverManager(driver_version=driver_version).install())
        return webdriver.Chrome(service=service, options=options)

    try:
        return build_driver(None)
    except SessionNotCreatedException as e:
        m = re.search(r"Current browser version is\s+(\d+)", str(e))
        if not m:
            raise
        major = m.group(1)
        return build_driver(major)


def wait_page_ready(driver, timeout=30):
    WebDriverWait(driver, timeout).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )

def dump_page(driver, html_name="page_dump.html", png_name="page_dump.png"):
    try:
        with open(html_name, "w", encoding="utf-8") as f:
            f.write(driver.page_source)
    except Exception:
        pass
    try:
        driver.save_screenshot(png_name)
    except Exception:
        pass

def ensure_inventory_page(driver, timeout=25):

    try:
        WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.ID, "inventory_container")))
        return
    except TimeoutException:
        driver.get("https://www.saucedemo.com/inventory.html")
        wait_page_ready(driver, timeout=30)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "inventory_container")))
        return


def login(driver, username, password):
    driver.set_page_load_timeout(45)
    attempts = 3
    last_error = None

    for i in range(1, attempts + 1):
        try:
            driver.get("https://www.saucedemo.com/")
            wait_page_ready(driver, timeout=30)

            user_el = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.ID, "user-name"))
            )
            pwd_el = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "password"))
            )
            btn_el = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "login-button"))
            )

            user_el.clear()
            user_el.send_keys(username)
            pwd_el.clear()
            pwd_el.send_keys(password)
            btn_el.click()

            ensure_inventory_page(driver, timeout=20)
            return

        except TimeoutException:
            try:
                err = WebDriverWait(driver, 2).until(
                    EC.visibility_of_element_located((By.CSS_SELECTOR, "[data-test='error']"))
                )
                if "locked out" in err.text.lower():
                    driver.get("https://www.saucedemo.com/")
                    wait_page_ready(driver, timeout=20)
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "user-name"))).send_keys("standard_user")
                    driver.find_element(By.ID, "password").send_keys("secret_sauce")
                    driver.find_element(By.ID, "login-button").click()
                    ensure_inventory_page(driver, timeout=20)
                    return
            except Exception:
                pass

            last_error = "Timeout на логине"
            dump_page(driver, f"saucedemo_login_fail_attempt{i}.html", f"saucedemo_login_fail_attempt{i}.png")
            time.sleep(2)
        except Exception as e:
            last_error = e
            dump_page(driver, f"saucedemo_login_fail_attempt{i}.html", f"saucedemo_login_fail_attempt{i}.png")
            time.sleep(2)

    raise TimeoutException(f"Не удалось авторизоваться. Последняя ошибка: {last_error}")


def find_sort_select(driver, timeout=25):
    wait = WebDriverWait(driver, timeout)

    try:
        el = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "[data-test='product_sort_container']")))
        return el
    except TimeoutException:
        pass

    try:
        el = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "product_sort_container")))
        return el
    except TimeoutException:
        pass

    try:
        for sel in ["select[data-test='product_sort_container']",
                    ".product_sort_container",
                    "select.product_sort_container"]:
            el = driver.execute_script("return document.querySelector(arguments[0]);", sel)
            if el:
                return el
    except Exception:
        pass

    dump_page(driver, "page_dump_before_sort.html", "before_sort.png")
    raise TimeoutException("Не найден элемент сортировки product_sort_container (см. before_sort.png и page_dump_before_sort.html)")


def apply_sort(driver, sort_value: str):
    """
    sort_value: 'az' | 'za' | 'lohi' | 'hilo'
    """
    ensure_inventory_page(driver, timeout=25)

    try:
        driver.execute_script("window.scrollTo(0, 0);")
    except Exception:
        pass

    select_el = find_sort_select(driver, timeout=25)

    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", select_el)
        time.sleep(0.2)
    except Exception:
        pass

    try:
        Select(select_el).select_by_value(sort_value)
    except (ElementNotInteractableException, ElementClickInterceptedException, NoSuchElementException, Exception):
        driver.execute_script(
            "const sel = arguments[0]; sel.value = arguments[1]; sel.dispatchEvent(new Event('change', {bubbles:true}));",
            select_el, sort_value
        )

    time.sleep(0.5)
    current_value = driver.execute_script("return arguments[0].value;", select_el)
    if current_value != sort_value:
        dump_page(driver, "page_dump_after_sort.html", "after_sort.png")
        raise RuntimeError(f"Сортировка не применилась (ожидали '{sort_value}', получили '{current_value}'). "
                           f"См. after_sort.png/page_dump_after_sort.html")


# ---------- Сбор карточек товаров ----------
def collect_items(driver) -> List[InventoryItem]:
    items = driver.find_elements(By.CSS_SELECTOR, ".inventory_item")
    result: List[InventoryItem] = []
    for it in items:
        name = it.find_element(By.CSS_SELECTOR, ".inventory_item_name").text.strip()
        price_text = it.find_element(By.CSS_SELECTOR, ".inventory_item_price").text.strip().replace("$", "")
        try:
            price = float(price_text)
        except ValueError:
            price = 0.0
        add_btn = it.find_element(By.CSS_SELECTOR, "button.btn_inventory")
        result.append(InventoryItem(it, name, price, add_btn))
    return result


# ---------- Выбор товара ----------
def pick_item(items: List[InventoryItem], strategy: str) -> InventoryItem:
    if not items:
        raise RuntimeError("Список товаров пуст.")

    s = strategy.lower()
    if s == "first":
        return items[0]
    if s == "last":
        return items[-1]
    if s == "cheapest":
        return min(items, key=lambda x: x.price)
    if s == "expensive":
        return max(items, key=lambda x: x.price)
    raise ValueError(f"Неизвестная стратегия выбора товара: {strategy}")


def do_cart_action(driver, action: str):
    wait = WebDriverWait(driver, 15)

    if action == "stop":
        return

    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a.shopping_cart_link"))).click()
    wait.until(EC.visibility_of_element_located((By.ID, "cart_contents_container")))

    if action == "remove":
        remove_btns = driver.find_elements(By.CSS_SELECTOR, "button.cart_button")
        if remove_btns:
            remove_btns[0].click()
        return

    if action == "continue":
        wait.until(EC.element_to_be_clickable((By.ID, "continue-shopping"))).click()
        wait.until(EC.visibility_of_element_located((By.ID, "inventory_container")))
        return

    if action == "checkout":
        wait.until(EC.element_to_be_clickable((By.ID, "checkout"))).click()
        wait.until(EC.visibility_of_element_located((By.ID, "first-name"))) 
        return

    raise ValueError(f"Неизвестное действие с корзиной: {action}")



def main():
    try:
        driver = start_opera_driver()
    except Exception as e:
        print(f"[FATAL] Не удалось запустить Opera: {e}", file=sys.stderr)
        sys.exit(1)

    try:
        login(driver, USERNAME, PASSWORD)

        ensure_inventory_page(driver, timeout=25)

        try:
            dump_page(driver, "page_dump_before_sort.html", "before_sort.png")
        except Exception:
            pass

        apply_sort(driver, SORT_VALUE)

        items = collect_items(driver)

        item = pick_item(items, PICK_STRATEGY)

        item.add_btn.click()

        do_cart_action(driver, CART_ACTION)

        time.sleep(2)
        print("[OK] Скрипт завершён без ошибок.")
    finally:
        pass


if __name__ == "__main__":
    main()
