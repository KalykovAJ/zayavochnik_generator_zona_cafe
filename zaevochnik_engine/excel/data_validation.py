from openpyxl.worksheet.datavalidation import DataValidation


def build_qty_validations(worksheet, validation_prompts: dict):
    """Создаёт и регистрирует на листе два правила Data Validation —
    для заказа в штуках и в упаковках. Возвращает (dv_unit, dv_pack)."""
    dv_unit = DataValidation(type="whole", operator="greaterThanOrEqual", formula1="0")
    dv_pack = DataValidation(type="whole", operator="greaterThanOrEqual", formula1="0")

    for key, dv in [("unit", dv_unit), ("pack", dv_pack)]:
        prompt = validation_prompts.get(key, {"title": "Внимание", "message": "Заполните поле"})
        dv.promptTitle = prompt["title"]
        dv.prompt = prompt["message"]
        dv.errorTitle = "Ошибка ввода"
        dv.error = "Введенное значение не соответствует правилам."
        dv.showInputMessage = True
        dv.showErrorMessage = True
        worksheet.add_data_validation(dv)

    return dv_unit, dv_pack
