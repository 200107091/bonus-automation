import re
import itertools
import pandas as pd
import openpyxl.styles
from openpyxl.utils import get_column_letter

bonus_path = r"C:\Users\52720\Dropbox\Shared with Assel\Расчет бонуса\Тендер\2026\2026_05\Бонусы.xlsx"

brands_path = r"C:\Users\52720\Dropbox\Shared with Assel\Расчет бонуса\Тендер\2026\2026_05\Брендс.xlsx"
bythim_path = r"C:\Users\52720\Dropbox\Shared with Assel\Расчет бонуса\Тендер\2026\2026_05\Бытхим.xlsx"

pk_path = r"C:\Users\52720\Dropbox\Shared with Assel\Расчет бонуса\Тендер\2026\2026_05\ПК.xlsx"
a17_path = r"C:\Users\52720\Dropbox\Shared with Assel\Расчет бонуса\Тендер\2026\2026_05\А17.xlsx"

sierra_path = r"C:\Users\52720\Dropbox\Shared with Assel\Расчет бонуса\Тендер\2026\2026_05\Сиерра.xlsx"

price_path = r"C:\Users\52720\Dropbox\Shared with Assel\Расчет бонуса\Тендер\2026\2026_05\Прайс.xlsx"
result_path = r"C:\Users\52720\Dropbox\Shared with Assel\Расчет бонуса\Тендер\2026\2026_05\bonus_ready_result.xlsx"

TOLERANCE = 2
MAX_COMBO_SIZE = 3


def clean_bin(x):
    if pd.isna(x):
        return ""

    try:
        if isinstance(x, (int, float)) and not pd.isna(x):
            text = str(int(x))
        else:
            text = str(x).strip()
            if "E+" in text.upper() or "E-" in text.upper():
                text = str(int(float(text)))
    except Exception:
        text = str(x).strip()

    text = text.replace(" ", "").replace(",", "").replace(".0", "")
    text = re.sub(r"\D", "", text)
    text = text.lstrip("0")
    return text

def clean_number(x):
    if pd.isna(x):
        return None

    text = str(x).strip()
    text = text.replace(" ", "")
    text = text.replace("\xa0", "")
    text = text.replace(",", ".")

    return pd.to_numeric(text, errors="coerce")


def clean_name(x):
    if pd.isna(x):
        return ""

    text = str(x).lower()
    text = text.replace("ё", "е")

    text = re.sub(r"[\"'«».,()\-]", " ", text)

    text = re.sub(
        r"\bтоо\b|\bип\b|\bкгу\b|\bгкп\b|\bгккп\b|\bкгкп\b|\bгу\b|\bао\b",
        " ",
        text
    )

    text = re.sub(r"\s+", " ", text).strip()

    return text



def normalize_pk_sales(df, source_name):
    df.columns = df.columns.astype(str).str.strip()

    df = df.rename(columns={
        "Склад": "Регион",
        "Документ движения (Регистратор).Ответственный": "Торговый менеджер",
        "Артикул Брендс": "Артикул",
        "Количество реализации": "Количество",
        "Сумма реализации": "Сумма со скидкой"
    })

    df["Источник"] = source_name

    if "БИН / ИИН" not in df.columns and "БИН" in df.columns:
        df["БИН / ИИН"] = df["БИН"]

    if "БИН / ИИН" in df.columns:
        df["БИН / ИИН"] = df["БИН / ИИН"].apply(clean_bin)
    else:
        df["БИН / ИИН"] = ""

    if "Партнер" not in df.columns:
        df["Партнер"] = df["Контрагент"]

    return df


def find_exact_or_combo(candidates, target):
    positive = candidates[candidates["shipment_remaining"] > 0].copy()
    positive = positive.sort_values("shipment_id")

    exact = positive[
        (positive["shipment_remaining"] - target).abs() <= TOLERANCE
    ]

    if not exact.empty:
        return [exact.sort_values("shipment_id").index[0]]

    positive["diff"] = (positive["shipment_remaining"] - target).abs()
    small = positive.sort_values("diff").head(50)

    idx_list = small.index.tolist()

    for size in range(2, min(MAX_COMBO_SIZE, len(idx_list)) + 1):
        for combo in itertools.combinations(idx_list, size):
            total = small.loc[list(combo), "shipment_remaining"].sum()

            if abs(total - target) <= TOLERANCE:
                return list(combo)

    return []

def prepare_payments(sheet_name):
    payments = pd.read_excel(bonus_path, sheet_name=sheet_name)
    payments.columns = payments.columns.astype(str).str.strip()

    payments = payments[
        [
            "№",
            "Выписка",
            "Город",
            "Наименование объекта",
            "БИН",
            "Дата оплаты",
            "Сумма",
            "Накладные расходы",
            "Чистая сумма",
            "Менеджер"
        ]
    ].copy()

    payments["Сумма"] = pd.to_numeric(payments["Сумма"], errors="coerce")
    payments["Дата оплаты"] = pd.to_datetime(payments["Дата оплаты"], errors="coerce")
    payments["Накладные расходы"] = pd.to_numeric(
        payments["Накладные расходы"],
        errors="coerce"
    ).fillna(0)

    payments["bin_clean"] = payments["БИН"].apply(clean_bin)
    payments["name_clean"] = payments["Наименование объекта"].apply(clean_name)
    payments = payments.dropna(subset=["Наименование объекта", "Сумма"])

    return payments


def prepare_sales_ab():
    brands = pd.read_excel(brands_path)
    bythim = pd.read_excel(bythim_path)

    brands["Источник"] = "Брендс"
    bythim["Источник"] = "Бытхим"

    sales = pd.concat([brands, bythim], ignore_index=True)
    return prepare_common_sales(sales)


def prepare_sales_pk():
    pk = pd.read_excel(pk_path)
    a17 = pd.read_excel(a17_path)

    pk = normalize_pk_sales(pk, "ПК")
    a17 = normalize_pk_sales(a17, "Аврора17")

    sales = pd.concat([pk, a17], ignore_index=True)
    return prepare_common_sales(sales)


def prepare_sales_sierra():
    sierra = pd.read_excel(sierra_path)
    sierra["Источник"] = "Сиерра"
    return prepare_common_sales(sierra)


def prepare_common_sales(sales):
    sales.columns = sales.columns.astype(str).str.strip()

    if "Сумма со скидкой" not in sales.columns:
        for alt in ["Сумма", "Сумма реализации", "Сумма отгрузки", "Сумма с НДС"]:
            if alt in sales.columns:
                sales["Сумма со скидкой"] = sales[alt]
                break

    if "Сумма со скидкой" not in sales.columns:
        raise KeyError(
            "Не найдена колонка 'Сумма со скидкой' в данных продаж. "
            "Проверьте заголовки в файле sales."
        )

    if "Цена" not in sales.columns:
        for alt in ["Цена за единицу", "Цена за ед.", "Цена за ед", "Цена с НДС"]:
            if alt in sales.columns:
                sales["Цена"] = sales[alt]
                break

    if "Цена" not in sales.columns:
        sales["Цена"] = None

    sales["Сумма со скидкой"] = sales["Сумма со скидкой"].apply(clean_number)
    sales["Цена"] = sales["Цена"].apply(clean_number)

    if "Количество" not in sales.columns:
        possible_qty_cols = [c for c in sales.columns if "Количество" in c]
        if possible_qty_cols:
            sales["Количество"] = sales[possible_qty_cols[0]].apply(clean_number)
        else:
            sales["Количество"] = None
    else:
        sales["Количество"] = sales["Количество"].apply(clean_number)

    # Егер цена бос болса, сумма / количество арқылы есептейміз
    sales["Цена"] = sales.apply(
        lambda r: r["Сумма со скидкой"] / r["Количество"]
        if (
            pd.isna(r["Цена"])
            and pd.notna(r["Сумма со скидкой"])
            and pd.notna(r["Количество"])
            and r["Количество"] != 0
        )
        else r["Цена"],
        axis=1
    )

    sales = sales[sales["Сумма со скидкой"] > 0].copy()

    if "Контрагент" not in sales.columns:
        sales["Контрагент"] = ""

    if "Партнер" not in sales.columns:
        sales["Партнер"] = sales["Контрагент"]

    if "Номенклатура" not in sales.columns:
        sales["Номенклатура"] = ""

    if "Артикул" not in sales.columns:
        sales["Артикул"] = ""

    if "БИН / ИИН" not in sales.columns and "БИН" in sales.columns:
        sales["БИН / ИИН"] = sales["БИН"].astype(str)
    elif "БИН / ИИН" in sales.columns and "БИН" in sales.columns:
        missing_bin = sales["БИН / ИИН"].isna() | (sales["БИН / ИИН"].astype(str).str.strip() == "")
        sales.loc[missing_bin, "БИН / ИИН"] = sales.loc[missing_bin, "БИН"].astype(str)

    if "БИН / ИИН" not in sales.columns:
        sales["БИН / ИИН"] = ""

    sales["bin_clean"] = sales["БИН / ИИН"].apply(clean_bin)

    sales["shipment_id"] = sales.index
    sales["shipment_remaining"] = sales["Сумма со скидкой"].copy()

    return sales


def add_bonus_row(result, status, method, pay, sale=None, allocated=0, remaining_payment=0, note=""):
    if sale is not None:
        price = sale["Цена"]
        original_sale_sum = sale["Сумма со скидкой"]
        original_qty = sale["Количество"]

        if pd.notna(price) and price != 0:
            qty_calc = allocated / price
        else:
           qty_calc = original_qty

     
        row = {
            "Статус": status,
            "Метод": method,
            "Комментарий": note,
            "№": pay["№"],
            "Регион": pay["Город"],
            "Менеджер": pay["Менеджер"],
            "Организация на которую заключается договор": pay["Выписка"],
            "Наименование Заказчика": pay["Наименование объекта"],
            "БИН": pay["БИН"],
            "Номенклатура": sale["Номенклатура"],
            "Артикул": sale["Артикул"],
            "Количество, ед.": qty_calc,
            "Цена за единицу победителя с НДС": price,
            "Итоговая сумма победителя с НДС": allocated,
            "Сумма для расчета бонуса": allocated,
            "Накладные": pay["Накладные расходы"],
            "Источник": sale["Источник"],
            "Контрагент из продаж": sale["Контрагент"],
            "Партнер из продаж": sale["Партнер"],
            "БИН / ИИН из продаж": sale.get("БИН / ИИН", sale.get("БИН", "")),
            "Сумма отгрузки": original_sale_sum,
            "Остаток оплаты": remaining_payment,
            "Остаток отгрузки после": sale["shipment_remaining"]
        }
    else:
        row = {
            "Статус": status,
            "Метод": method,
            "Комментарий": note,
            "№": pay["№"],
            "Регион": pay["Город"],
            "Менеджер": pay["Менеджер"],
            "Организация на которую заключается договор": pay["Выписка"],
            "Наименование Заказчика": pay["Наименование объекта"],
            "БИН": pay["БИН"],
            "Номенклатура": "",
            "Артикул": "",
            "Количество, ед.": "",
            "Цена за единицу победителя с НДС": "",
            "Итоговая сумма победителя с НДС": "",
            "Сумма для расчета бонуса": "",
            "Накладные": pay["Накладные расходы"],
            "Источник": "",
            "Контрагент из продаж": "",
            "Партнер из продаж": "",
            "БИН / ИИН из продаж": "",
            "Сумма отгрузки": "",
            "Остаток оплаты": remaining_payment,
            "Остаток отгрузки после": ""
        }

    result.append(row)


def allocate_payments(payments, sales):
    result = []
    unmatched_payments = []

    # PASS 1 — exact / combo
    for _, pay in payments.iterrows():
        payment_remaining = pay["Сумма"]

        if pay["bin_clean"] == "":
            add_bonus_row(
                result,
                "Не найдено",
                "Нет БИН в файле бонусов",
                pay,
                remaining_payment=payment_remaining,
                note="Невозможно сопоставить без БИН"
            )
            continue

        client_shipments = sales[
    (sales["bin_clean"] == pay["bin_clean"]) &
    (sales["shipment_remaining"] > 0)
].copy()  

        if client_shipments.empty:
            unmatched_payments.append(pay)
            continue

        combo_indexes = find_exact_or_combo(client_shipments, payment_remaining)

        if combo_indexes:
            for idx in combo_indexes:
                if payment_remaining <= TOLERANCE:
                    break

                sale = sales.loc[idx]
                available = sale["shipment_remaining"]

                if available <= 0:
                    continue

                allocated = min(payment_remaining, available)

                sales.loc[idx, "shipment_remaining"] = available - allocated
                payment_remaining -= allocated

                if abs(payment_remaining) <= TOLERANCE:
                    payment_remaining = 0

                sale_after = sales.loc[idx].copy()

                add_bonus_row(
                    result,
                    "Сопоставлено",
                    "Точное совпадение суммы",
                    pay,
                    sale=sale_after,
                    allocated=allocated,
                    remaining_payment=payment_remaining,
                    note="Оплата закрыта точной суммой одной или нескольких отгрузок"
                )
        else:
            unmatched_payments.append(pay)

    # PASS 2 — FIFO
    for pay in unmatched_payments:
        payment_remaining = pay["Сумма"]

        client_shipments = sales[
    (sales["bin_clean"] == pay["bin_clean"]) &
    (sales["shipment_remaining"] > 0)
].sort_values("shipment_id")

        if client_shipments.empty:
            add_bonus_row(
                result,
                "Не найдено",
                "БИН есть, но отгрузки не найдены",
                pay,
                remaining_payment=payment_remaining,
                note="По данному БИН нет доступных отгрузок"
            )
            continue

        for idx, sale in client_shipments.iterrows():
            if payment_remaining <= 0:
                break

            available = sale["shipment_remaining"]
            allocated = min(payment_remaining, available)

            sales.loc[idx, "shipment_remaining"] = available - allocated
            payment_remaining -= allocated

            if abs(payment_remaining) <= TOLERANCE:
                payment_remaining = 0

            sale_after = sales.loc[idx].copy()

            if allocated < available:
                status = "Сопоставлено частично"
                note = "Оплата закрывает только часть отгрузки"
            else:
                status = "Сопоставлено FIFO"
                note = "Оплата закрыта по FIFO: самые старые отгрузки"

            add_bonus_row(
                result,
                status,
                "FIFO: самые старые отгрузки",
                pay,
                sale=sale_after,
                allocated=allocated,
                remaining_payment=payment_remaining,
                note=note
            )

        if payment_remaining > 0:
            add_bonus_row(
                result,
                "Оплата закрыта не полностью",
                "FIFO: самые старые отгрузки",
                pay,
                remaining_payment=payment_remaining,
                note="Суммы доступных отгрузок не хватило для полного закрытия оплаты"
            )

    return pd.DataFrame(result)


def add_price_and_formulas(result_df):
    price = pd.read_excel(price_path)
    price.columns = price.columns.astype(str).str.strip()

    price = price[["Номенклатура", "Прайс"]].copy()
    price["Прайс"] = pd.to_numeric(price["Прайс"], errors="coerce")

    price["Номенклатура"] = price["Номенклатура"].astype(str).str.strip()

    for col in ["Номенклатура", "Накладные", "Сумма для расчета бонуса", "Цена за единицу победителя с НДС"]:
        if col not in result_df.columns:
            result_df[col] = ""

    result_df["Номенклатура"] = result_df["Номенклатура"].astype(str).str.strip()

    price = price.drop_duplicates(subset=["Номенклатура"], keep="first")

    if "Прайс" in result_df.columns:
        result_df = result_df.drop(columns=["Прайс"])

    result_df = result_df.merge(
        price,
        on="Номенклатура",
        how="left",
        validate="many_to_one"
    )

    result_df["Прайс"] = pd.to_numeric(result_df["Прайс"], errors="coerce")
    result_df["Накладные"] = pd.to_numeric(result_df["Накладные"], errors="coerce").fillna(0)

    result_df["Сумма для расчета бонуса"] = pd.to_numeric(
        result_df["Сумма для расчета бонуса"],
        errors="coerce"
    ).fillna(0)

    result_df["Цена за единицу победителя с НДС"] = pd.to_numeric(
        result_df["Цена за единицу победителя с НДС"],
        errors="coerce"
    )

    result_df["Потери"] = (
        result_df["Сумма для расчета бонуса"]
        * result_df["Накладные"]
        / 100
        * 1.11
    )

    result_df["Сумма чистыми"] = result_df["Сумма для расчета бонуса"] - result_df["Потери"]

    result_df["Разница"] = (
        result_df["Цена за единицу победителя с НДС"] - result_df["Прайс"]
    )

    def calc_bonus_percent(x):
     if pd.isna(x):
        return ""

    # Егер разница -100 мен 0 арасында болса: -100, -86, -15, -2, -1, 0
     if -100 <= x <= 0:
        return 0.05

    # Егер разница плюспен шықса
     if x > 0:
        return 0.05

    # Егер -100-ден төмен болса: -101, -200, -450
     return 0

    result_df["% бонус"] = result_df["Разница"].apply(calc_bonus_percent)

    result_df["bonus_percent_num"] = pd.to_numeric(result_df["% бонус"], errors="coerce").fillna(0)

    result_df["Бонус"] = (
     result_df["Сумма чистыми"]
    * result_df["bonus_percent_num"]
    * 0.89
)

    result_df["рук"] = result_df.apply(
     lambda row: row["Сумма чистыми"] * 0.02 * 0.89
     if row["bonus_percent_num"] == 0.05
     else 0,
     axis=1
)

    result_df = result_df.drop(columns=["bonus_percent_num"])
    
    final_columns = [
        "Статус",
        "Метод",
        "Комментарий",
        "№",
        "Регион",
        "Менеджер",
        "Организация на которую заключается договор",
        "Наименование Заказчика",
        "БИН",
        "Номенклатура",
        "Артикул",
        "Количество, ед.",
        "Цена за единицу победителя с НДС",
        "Итоговая сумма победителя с НДС",
        "Сумма для расчета бонуса",
        "Накладные",
        "Потери",
        "Сумма чистыми",
        "Прайс",
        "Разница",
        "% бонус",
        "Бонус",
        "рук",
        "Комментарий расчет",
        "Источник",
        "Контрагент из продаж",
        "Партнер из продаж",
        "БИН / ИИН из продаж",
        "Сумма отгрузки",
        "Остаток оплаты",
        "Остаток отгрузки после"
    ]

    for col in final_columns:
        if col not in result_df.columns:
            result_df[col] = ""

    return result_df[final_columns]

# -----------------------------------------------------------------------------

def make_summary(df, block_name):
    work = df.copy()

    work["Сумма чистыми"] = pd.to_numeric(work["Сумма чистыми"], errors="coerce").fillna(0)
    work["Бонус"] = pd.to_numeric(work["Бонус"], errors="coerce").fillna(0)
    work["рук"] = pd.to_numeric(work["рук"], errors="coerce").fillna(0)

    summary = work.groupby(
        ["Регион", "Менеджер"],
        as_index=False
    ).agg({
        "Сумма чистыми": "sum",
        "Бонус": "sum",
        "рук": "sum"
    })

    summary["Итого бонус"] = summary["Бонус"] + summary["рук"]
    summary["Блок"] = block_name

    return summary

# -----------------------------------------------------------------------------

def norm_name(x):
    if pd.isna(x):
        return ""

    s = str(x).lower().strip()
    s = re.sub(r"[^а-яёa-z ]", " ", s)
    words = [w for w in s.split() if w]

    aliases = {
        "гулжайна": "жайна",
        "гүлжайна": "гулжайна",
        "айнагуль": "айнагул",
        "айнагүл": "айнагул",
        "сансызбеккызы": "сансызбекқызы",
    }

    words = [aliases.get(w, w) for w in words]
    words = sorted(words)  # ең маңызды жол

    return " ".join(words)


def make_final_bonus_summary(summary_all, result_ab, result_pk):
    df = summary_all.copy()

    new_calc = make_new_bonus_calc(result_ab, result_pk)

    excluded_keys = set(new_calc["name_key"].dropna().unique())

    df["name_key"] = df["Менеджер"].apply(norm_name)

# Новый расчетта бар менеджерлерді сводтағы бонус/рук есептен алып тастаймыз
    df_for_ruk = df[~df["name_key"].isin(excluded_keys)].copy()

    for col in ["Сумма чистыми", "Бонус", "рук"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    df["name_key"] = df["Менеджер"].apply(norm_name)

    # 1. Личный бонус менеджера
    personal = df_for_ruk.groupby(["Менеджер", "name_key"], as_index=False).agg({
        "Сумма чистыми": "sum",
        "Бонус": "sum",
        "рук": "sum"
    })

    personal = personal.rename(columns={
        "Бонус": "Бонус личный",
        "рук": "Рук начислено в строках"
    })
    personal["Бонус личный"] = pd.to_numeric(personal["Бонус личный"], errors="coerce").fillna(0).astype(float)
    personal["Рук начислено в строках"] = pd.to_numeric(personal["Рук начислено в строках"], errors="coerce").fillna(0).astype(float)
    personal["Сумма чистыми"] = pd.to_numeric(personal["Сумма чистыми"], errors="coerce").fillna(0).astype(float)

    personal["Бонус руководителя"] = 0.0

    # 2. Иерархия руководителей
    hierarchy = {
        norm_name("Казанцев Сергей"): [
        ],
        norm_name("Малик Дулат"): [
            norm_name("Сансызбекқызы Мақпал"),
            norm_name("Назира Қожаева"),
            norm_name("Кадирханов Ерлан"),
            norm_name("Акерке Камза"),
            norm_name("Рауан Турганбай"),
        ],
        norm_name("Котельников Сергей"): [
            norm_name("Дентаев Данияр"),
            norm_name("Айнагул Дарбаева"),
            norm_name("Лаура Султанбекова"),
        ],
        norm_name("Гулжайна Серикбай"): [
            norm_name("Малика Нуршина"),
        ],
        norm_name("Руслан Ли"): [
            norm_name("Иманбек Сыздықов"),
            norm_name("Турсынай Зикенов"),
        ],
    }

    for boss_key, employee_keys in hierarchy.items():
        ruk_sum = personal.loc[
            personal["name_key"].isin(employee_keys),
            "Рук начислено в строках"
        ].sum()

        personal.loc[
            personal["name_key"] == boss_key,
            "Бонус руководителя"
        ] += ruk_sum

    # 3. Дополнительный блок 0,5%
    total_clean = personal["Сумма чистыми"].sum()
    extra_total = total_clean * 0.005 * 0.89


    mandatory_people = [
        "Гулжайна Серикбай",
        "Малика Нуршина",
        "Анастасия Доля"
    ]

    for person in mandatory_people:
        key = norm_name(person)

        if key not in personal["name_key"].values:
            personal = pd.concat([
                personal,
                pd.DataFrame([{
                    "Менеджер": person,
                    "name_key": key,
                    "Сумма чистыми": 0.0,
                    "Бонус личный": 0.0,
                    "Рук начислено в строках": 0.0,
                    "Бонус руководителя": 0.0,
                    "Доп бонус 0.5%": 0.0
                }])
            ], ignore_index=True)

    personal["Доп бонус 0.5%"] = 0.0

    zhayna_key = norm_name("Гулжайна Серикбай")
    malika_key = norm_name("Малика Нуршина")
    anastasia_key = norm_name("Анастасия Доля")

    personal.loc[personal["name_key"] == zhayna_key, "Доп бонус 0.5%"] += extra_total * 0.60
    personal.loc[personal["name_key"] == malika_key, "Доп бонус 0.5%"] += extra_total * 0.20
    personal.loc[personal["name_key"] == anastasia_key, "Доп бонус 0.5%"] += extra_total * 0.20

    personal["Итого бонус"] = (
        personal["Бонус личный"]
        + personal["Бонус руководителя"]
        + personal["Доп бонус 0.5%"]
    )


    # Новый расчеттағы адамдарды итогқа міндетті түрде қосамыз
    new_calc = make_new_bonus_calc(result_ab, result_pk)

    new_part = new_calc[[
      "ФИО",
      "name_key",
      "Итог бонус за деньги",
      "Итог бонус за договор",
      "Итого бонус"
   ]].copy()

    new_part = new_part.rename(columns={
      "ФИО": "Менеджер",
      "Итог бонус за деньги": "Бонус план деньги",
      "Итог бонус за договор": "Бонус план договор",
      "Итого бонус": "Бонус новый расчет"
  })

    for col in ["Бонус личный", "Рук начислено в строках", "Бонус руководителя", "Доп бонус 0.5%"]:
     new_part[col] = 0.0

     new_part["Сумма чистыми"] = 0.0

     new_part["Итого бонус"] = new_part["Бонус новый расчет"]

# personal жақта жоқ колонкаларды қосамыз
    for col in new_part.columns:
     if col not in personal.columns:
        personal[col] = 0

    for col in personal.columns:
     if col not in new_part.columns:
        new_part[col] = 0

    personal = pd.concat([personal, new_part[personal.columns]], ignore_index=True)

# Бір адам екі рет шықса, name_key бойынша біріктіреміз
    personal = personal.groupby(["name_key"], as_index=False).agg({
      "Менеджер": "first",
      "Сумма чистыми": "sum",
      "Бонус личный": "sum",
      "Рук начислено в строках": "sum",
      "Бонус руководителя": "sum",
      "Доп бонус 0.5%": "sum",
      "Бонус план деньги": "sum",
      "Бонус план договор": "sum",
      "Бонус новый расчет": "sum",
      "Итого бонус": "sum"
  })

    return personal


# -----------------------------------------------------------------------------
def norm_name(x):
    if pd.isna(x):
        return ""

    text = str(x).lower().replace("ё", "е")
    text = re.sub(r"[^а-яa-z\s]", " ", text)
    words = text.split()

    # фамилия/имя орын ауысса да бірдей болу үшін
    words = sorted(words)

    return " ".join(words)


def find_col(df, words):
    for col in df.columns:
        name = str(col).lower()
        if all(w.lower() in name for w in words):
            return col
    return None


def read_itog_plan():
    raw = pd.read_excel(bonus_path, sheet_name="ИТОГ", header=None)

    header_row = raw[
        raw.apply(
            lambda r: r.astype(str).str.contains("ФИО", case=False, na=False).any(),
            axis=1
        )
    ].index[0]

    plan = pd.read_excel(bonus_path, sheet_name="ИТОГ", header=header_row)
    plan.columns = plan.columns.astype(str).str.strip()

    fio_col = find_col(plan, ["ФИО"])

    result = pd.DataFrame()
    result["ФИО"] = plan[fio_col]
    result["name_key"] = result["ФИО"].apply(norm_name)

    result["План деньги"] = pd.to_numeric(plan[find_col(plan, ["План", "деньги"])], errors="coerce")
    result["План договор"] = pd.to_numeric(plan[find_col(plan, ["План", "договор"])], errors="coerce")
    result["Факт договора"] = pd.to_numeric(plan[find_col(plan, ["Факт", "договор"])], errors="coerce")
    result["Ставка"] = pd.to_numeric(plan[find_col(plan, ["Ставка"])], errors="coerce")
    result["Бонус за договор"] = pd.to_numeric(plan[find_col(plan, ["Бонус", "договор"])], errors="coerce")
    result["Бонус за деньги"] = pd.to_numeric(plan[find_col(plan, ["Бонус", "деньги"])], errors="coerce")

    result = result.dropna(subset=["ФИО"])
    return result


def calc_contract_bonus(row):
    plan = row["План договор"]
    fact = row["Факт договора"]
    base_bonus = row["Бонус за договор"]
    rate = row["Ставка"]

    if pd.isna(plan) or plan == 0 or pd.isna(fact):
        return 0

    pct = fact / plan * 100

    if pct >= 100:
        main_bonus = base_bonus
    elif pct >= 80:
        main_bonus = base_bonus * pct / 100
    elif pct >= 60:
        main_bonus = base_bonus * (pct / 100) / 2
    else:
        main_bonus = 0

    over_bonus = 0
    if pct > 100:
        over_bonus += min(pct - 100, 20) / 100 * rate
    if pct > 120:
        over_bonus += (pct - 120) * (rate * 0.005)

    return main_bonus + over_bonus


def calc_money_bonus(row):
    plan = row["План деньги"]
    fact = row["Факт деньги"]
    base_bonus = row["Бонус за деньги"]

    if pd.isna(plan) or plan == 0 or pd.isna(fact):
        return 0

    pct = fact / plan * 100

    if pct >= 100:
        return base_bonus
    elif pct >= 80:
        return base_bonus * pct / 100
    elif pct >= 60:
        return base_bonus * 0.30
    else:
        return 0


def make_new_bonus_calc(result_ab, result_pk):
    plan = read_itog_plan()

    money_source = pd.concat([result_ab, result_pk], ignore_index=True).copy()
    money_source["name_key"] = money_source["Менеджер"].apply(norm_name)
    money_source["Сумма для расчета бонуса"] = pd.to_numeric(
        money_source["Сумма для расчета бонуса"],
        errors="coerce"
    ).fillna(0)

    fact_money = money_source.groupby("name_key", as_index=False)["Сумма для расчета бонуса"].sum()
    fact_money = fact_money.rename(columns={"Сумма для расчета бонуса": "Факт деньги"})

    plan = plan.merge(fact_money, on="name_key", how="left")
    plan["Факт деньги"] = plan["Факт деньги"].fillna(0)

    # Сойнова Инна = өзінің + Митянина Екатерина + Аниров Темирлан
    soy_key = norm_name("Сойнова Инна")
    extra_keys = [
        norm_name("Митянина Екатерина"),
        norm_name("Аниров Темирлан")
    ]

    extra_sum = plan.loc[plan["name_key"].isin(extra_keys), "Факт деньги"].sum()
    plan.loc[plan["name_key"] == soy_key, "Факт деньги"] += extra_sum

    plan["% деньги"] = plan["Факт деньги"] / plan["План деньги"]
    plan["% договора"] = plan["Факт договора"] / plan["План договор"]

    plan["Итог бонус за деньги"] = plan.apply(calc_money_bonus, axis=1)
    plan["Итог бонус за договор"] = plan.apply(calc_contract_bonus, axis=1)

    plan["Итого бонус"] = (
        plan["Итог бонус за деньги"].fillna(0)
        + plan["Итог бонус за договор"].fillna(0)
    )

    return plan

# -----------------------------------------------------------------------------

def style_sheet(ws):
    header_fill = openpyxl.styles.PatternFill("solid", fgColor="F4CCCC")
    total_fill = openpyxl.styles.PatternFill("solid", fgColor="FCE4D6")
    thin = openpyxl.styles.Side(style="thin", color="D9D9D9")
    border = openpyxl.styles.Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.freeze_panes = "A2"

    for cell in ws[1]:
        cell.font = openpyxl.styles.Font(bold=True)
        cell.fill = header_fill
        cell.alignment = openpyxl.styles.Alignment(vertical="center", wrap_text=True)
        cell.border = border

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.border = border
            cell.alignment = openpyxl.styles.Alignment(vertical="center", wrap_text=True)

            if isinstance(cell.value, (int, float)):
                cell.number_format = '#,##0.00'

    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)

        for cell in col:
            if cell.value is not None:
                max_len = max(max_len, len(str(cell.value)))

        ws.column_dimensions[col_letter].width = min(max_len + 2, 35)

        if cell.value in [
        "Наименование Заказчика",
           "Организация на которую заключается договор"
]:
         ws.column_dimensions[col_letter].width = 35

    ws.auto_filter.ref = ws.dimensions

    text_columns = ["№", "БИН", "БИН / ИИН из продаж", "Артикул"]
    integer_columns = ["Накладные"]

    header_map = {}
    for cell in ws[1]:
        header_map[str(cell.value).strip()] = cell.column

    for col_name in text_columns:
        if col_name in header_map:
            col_letter = get_column_letter(header_map[col_name])
            for cell in ws[col_letter][1:]:
                cell.number_format = "@"

    for col_name in integer_columns:
        if col_name in header_map:
            col_letter = get_column_letter(header_map[col_name])
            for cell in ws[col_letter][1:]:
                cell.number_format = "0"

    for row in ws.iter_rows():
     ws.row_dimensions[row[0].row].height = 18  

     wide_cols = {
       "A": 18,   # Статус
       "B": 25,   # Метод
       "C": 35,   # Комментарий
       "H": 40,   # Наименование заказчика
       "I": 18,   # БИН
       "J": 25,   # Номенклатура
       "N": 18,   # Итоговая сумма
       "Q": 18,   # Сумма чистыми
}

    for col, width in wide_cols.items():
     ws.column_dimensions[col].width = width          


# ===== ЗАПУСК =====

payments_ab = prepare_payments("АБ")
sales_ab = prepare_sales_ab()
result_ab = allocate_payments(payments_ab, sales_ab)
result_ab = add_price_and_formulas(result_ab)


payments_pk = prepare_payments("ПК")
sales_pk = prepare_sales_pk()
result_pk = allocate_payments(payments_pk, sales_pk)
result_pk = add_price_and_formulas(result_pk)

payments_sd = prepare_payments("СД")
sales_sd = prepare_sales_sierra()
result_sd = allocate_payments(payments_sd, sales_sd)
result_sd = add_price_and_formulas(result_sd)


summary_ab = make_summary(result_ab, "АБ")
summary_pk = make_summary(result_pk, "ПК")
summary_sd = make_summary(result_sd, "СД")

summary_all = pd.concat([summary_ab, summary_pk, summary_sd], ignore_index=True)

summary_total = summary_all.groupby(
    ["Регион", "Менеджер"],
    as_index=False
).agg({
    "Сумма чистыми": "sum",
    "Бонус": "sum",
    "рук": "sum",
    "Итого бонус": "sum"
})

new_bonus_calc = make_new_bonus_calc(result_ab, result_pk)
final_bonus_summary = make_final_bonus_summary(summary_all, result_ab, result_pk)

with pd.ExcelWriter(result_path, engine="openpyxl") as writer:
    result_ab.to_excel(writer, sheet_name="Готовый АБ", index=False)
    result_pk.to_excel(writer, sheet_name="Готовый ПК", index=False)
    result_sd.to_excel(writer, sheet_name="Готовый СД", index=False)


    problem_price_rows = pd.concat(
     [result_ab, result_pk, result_sd],
     ignore_index=True
     ).copy()

    problem_price_rows["Разница"] = pd.to_numeric(
    problem_price_rows["Разница"],
    errors="coerce"
)

    problem_price_rows = problem_price_rows[
     problem_price_rows["Разница"] < -100
     ].copy()

    problem_price_rows.to_excel(
      writer,
      sheet_name="Разница меньше -100",
      index=False
)
    summary_all.to_excel(writer, sheet_name="Свод", index=False)
    summary_total.to_excel(writer, sheet_name="Свод итог", index=False)
    new_bonus_calc.to_excel(writer, sheet_name="Новый расчет Python", index=False)
    final_bonus_summary.to_excel(writer, sheet_name="Итог бонусов", index=False)


    for ws in writer.sheets.values():

    # барлық колонкаға перенос өшіру
     for row in ws.iter_rows():
        for cell in row:
            cell.alignment = openpyxl.styles.Alignment(
                horizontal="left",
                vertical="center",
                wrap_text=False
            )

    # ұзын тексті автоматты анықтау
    for col in ws.iter_cols():
        max_len = max(
            len(str(c.value)) if c.value is not None else 0
            for c in col
        )

        if max_len > 40:
            for cell in col:
                cell.alignment = openpyxl.styles.Alignment(
                    horizontal="left",
                    vertical="center",
                    wrap_text=True
                )

    for ws in writer.book.worksheets:
        style_sheet(ws)




print("Готово:", result_path)
print("АБ:")
print(result_ab["Статус"].value_counts())
print("ПК:")
print(result_pk["Статус"].value_counts())
print("СД:")
print(result_sd["Статус"].value_counts())