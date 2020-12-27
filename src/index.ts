import ItemResponse = GoogleAppsScript.Forms.ItemResponse;
import GoogleSpreadsheet = GoogleAppsScript.Spreadsheet.Spreadsheet;
import GoogleSheet = GoogleAppsScript.Spreadsheet.Sheet;

class FormResponse {
    static parseFormResponse(itemResponses: ItemResponse[]): FormResponse {
        let dateString: string | undefined = null

        for (const itemResponse of itemResponses) {
            const item = itemResponse.getItem()
            switch (item.getTitle()) {
                case '請求月':
                    dateString = itemResponse.getResponse() as string
            }
        }

        const [year, month] = dateString.split('-').map((s) => Number(s))
        const date = new Date(year, month - 1)

        return new FormResponse(date)
    }

    constructor(readonly date: Date) {
    }
}

const spreadsheetDirections = [
    'next',
    'previous',
    'up',
    'down'
] as const

type SpreadsheetDirection = typeof spreadsheetDirections[number]

class SpreadsheetRange {
    constructor(readonly sheet: Sheet, readonly range: GoogleAppsScript.Spreadsheet.Range) {
    }

    get value(): unknown {
        return this.range.getValue()
    }

    set value(newValue) {
        this.range.setValue(newValue)
    }

    get row(): number {
        return this.range.getRow()
    }

    get column(): number {
        return this.range.getColumn()
    }

    nextCell(direction: SpreadsheetDirection): SpreadsheetRange {
        switch (direction) {
            case "next":
                return this.sheet.getRange(this.row, this.column + 1)
            case "previous":
                return this.sheet.getRange(this.row, this.column - 1)
            case "up":
                return this.sheet.getRange(this.row - 1, this.column)
            case "down":
                return this.sheet.getRange(this.row + 1, this.column)
        }
    }
}

class Spreadsheet {
    static make(id: string): Spreadsheet {
        return new Spreadsheet(SpreadsheetApp.openById(id))
    }

    constructor(readonly spreadsheet: GoogleSpreadsheet) {
    }

    getSheet(name: string): Sheet | undefined {
        const sheet = this.spreadsheet.getSheetByName(name)
        if (!sheet) {
            return undefined
        }

        return new Sheet(this, sheet)
    }
}

class Sheet {
    constructor(private readonly spreadsheet: Spreadsheet, private readonly sheet: GoogleSheet) {
    }

    findText(text: string): SpreadsheetRange | undefined {
        const range = this.sheet.createTextFinder(text).findNext();

        if (!range) {
            return undefined
        }

        return new SpreadsheetRange(this, range)
    }

    getRange(row: number, column: number): SpreadsheetRange {
        return new SpreadsheetRange(this, this.sheet.getRange(`${convertColumnIndexToNotation(column)}${row}`))
    }

    copy(name: string): Sheet {
        const newSheet = this.sheet.copyTo(this.spreadsheet.spreadsheet)
        newSheet.setName(name)
        return new Sheet(this.spreadsheet, newSheet)
    }

    moveTo(position: number) {
        this.spreadsheet.spreadsheet.setActiveSheet(this.sheet)
        this.spreadsheet.spreadsheet.moveActiveSheet(position)
    }
}

function convertNotationToIndex(column: string): number {
    const a = 'A'.charCodeAt(0)

    let output = 0
    for (let i = 0; i < column.length; i++) {
        const nextChar = column.charAt(i)
        const shift = 26 * i

        output += shift + (nextChar.charCodeAt(0) - a)
    }

    return output
}

function convertColumnIndexToNotation(column: number): string {
    let output = ''
    let temp = 0
    while (column > 0) {
        temp = (column - 1) % 26
        output = String.fromCharCode(temp + 65) + output
        column = (column - temp - 1) / 26
    }
    return output
}

class WorkDetail {
    static makeFromSheetName(name: string): WorkDetail | undefined {
        const spreadsheet = Spreadsheet.make(PropertiesService.getScriptProperties().getProperty('work_detail_id'))
        const sheet = spreadsheet.getSheet(name)
        if (!sheet) {
            return undefined
        }

        const itemHeaderRange = sheet.findText('項目名')
        if (!itemHeaderRange) {
            return undefined
        }

        let items: Item[] = []
        let currentCell = itemHeaderRange
        while (true) {
            currentCell = currentCell.nextCell("down")
            const item = Item.fromCell(currentCell)
            if (item === undefined) {
                continue
            }

            items.push(item)
            if (item.name === "工程管理") {
                break
            }
        }

        let totalTextCell = sheet.findText("合計")
        if (!totalTextCell) {
            return undefined
        }

        currentCell = totalTextCell.nextCell("next")
        const total = currentCell.value as number
        currentCell = currentCell.nextCell("up")
        const tax = currentCell.value as number
        currentCell = currentCell.nextCell("up")
        const subtotal = currentCell.value as number

        return new WorkDetail(name, items, subtotal, tax, total)
    }

    constructor(readonly name: string, readonly items: Item[], readonly subtotal: number, readonly tax: number, readonly total: number) {
    }
}

class Item {
    static fromCell(cell: SpreadsheetRange): Item {
        let currentCell = cell

        const name = currentCell.value as string
        if (name === undefined || name === '') {
            return undefined
        }

        currentCell = currentCell.nextCell("next")
        const requiredUnit = currentCell.value as number

        currentCell = currentCell.nextCell("next")
            .nextCell("next")
        const amount = currentCell.value as number
        currentCell = currentCell.nextCell("next")
        const unit = currentCell.value as string
        currentCell = currentCell.nextCell("next")
        const unitPrice = currentCell.value as number
        currentCell = currentCell.nextCell("next")
        const subtotal = currentCell.value as number

        return new Item(name, requiredUnit, unitPrice, amount, unit, subtotal)
    }

    constructor(
        readonly name: string, readonly requiredUnit: number, readonly unitPrice: number,
        readonly amount: number, readonly unit: string, readonly subtotal: number
    ) {
    }
}

class Invoice {
    constructor(readonly workDetail: WorkDetail) {
    }

    make() {
        const date = new Date()
        const spreadsheet = Spreadsheet.make(PropertiesService.getScriptProperties().getProperty('invoice_id'))
        const templateSheet = spreadsheet.getSheet('テンプレート')
        if(spreadsheet.getSheet(this.workDetail.name)) {
            throw new Error(`Sheet name ${this.workDetail.name} is already exist.`)
        }
        const sheet = templateSheet.copy(this.workDetail.name)
        sheet.moveTo(2)

        const invoiceDayCell = sheet.findText('請求日: ').nextCell("next")
        invoiceDayCell.value = Utilities.formatDate(date, 'JST', 'yyyy/MM/dd')
        const invoiceNumberCell = sheet.findText('請求番号: ').nextCell("next")
        invoiceNumberCell.value = Utilities.formatDate(date, 'JST', 'yyyyMMdd-01')

        this.setItems(sheet, this.workDetail.items)

        this.setRemarks(sheet, date)
    }

    private setItems(sheet: Sheet, items: Item[]) {
        let itemCell = sheet.findText('品 番 • 品 名')
        let amountCell = sheet.findText('数 量')
        let unitPriceCell = sheet.findText('単 価')
        for (const item of items) {
            itemCell = itemCell.nextCell('down')
            if (item.name === '工程管理') {
                itemCell = itemCell.nextCell("down")
                amountCell = amountCell.nextCell("down")
                unitPriceCell = unitPriceCell.nextCell("down")
            }
            itemCell.value = item.name
            amountCell = amountCell.nextCell("down")
            amountCell.value = item.amount
            let unitCell = amountCell.nextCell("next")
            unitCell.value = item.unit
            unitPriceCell = unitPriceCell.nextCell("down")
            unitPriceCell.value = item.unitPrice
        }
    }

    private setRemarks(sheet: Sheet, date: Date) {
        const limited = new Date(date.getFullYear(), date.getMonth() + 1, 0)

        const remarksCell = sheet.findText('備考').nextCell("down")
        remarksCell.value = remarksCell.value + Utilities.formatDate(limited, 'JST', 'M/dd')
    }
}

function submit(event: any) {
    const form = FormResponse.parseFormResponse(event.response.getItemResponses())
    const formatDateString = Utilities.formatDate(form.date, "JST", "yyyy/M")
    const workDetail = WorkDetail.makeFromSheetName(`${formatDateString}月分`)
    if (!workDetail) {
        console.error('sheet not found.')
        return
    }

    const invoice = new Invoice(workDetail)
    invoice.make()
}
