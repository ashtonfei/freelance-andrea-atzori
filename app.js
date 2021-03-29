const APP_NAME = "Email"
const RN_RECIPIENT = "recipient" // range name of recipient
const RN_SUBJECT = "subject" // range name of subject
const RN_CC = "cc" // range name of cc
const RN_BCC = "bcc" // range name of bcc
const RN_WIDTH = "width" // range name of width
const RN_TRIGGER = "trigger" // range name of trigger
const RN_BLNAK_ROW_HEIGHT = "blankRowHeight" // range name of trigger
const RN_MIN_COLUMN_WIDTH = "minColumnWidth" // range name of min column width
const RN_MAX_COLUMN_WIDTH = "maxColumnWidth" // range name of max column width
const RN_COLUMNS = "columns" // range name of column percentages
const RN_BG_COLOR = "bgColor" // range name of background color
const RN_TEXT_COLOR = "textColor" // range name of text color
const RN_BORDER_COLOR = "borderColor" // range name of border color
const RN_SHEET_URL_OF_CHARTS = "sheetUrlOfCharts" // range name of sheet url for charts


const RN_CONTENT = "body" // range name of email content

const SN_EMAIL = "Email Template" // sheet name of email content

function sendEmail() {
    const app = new App()
    app.sendEmail()
}

function onOpen() {
    const ui = SpreadsheetApp.getUi()
    const menu = ui.createMenu(APP_NAME)
    menu.addItem("Send", "sendEmail")
    menu.addToUi()
}

class App {
    constructor() {
        this.ss = SpreadsheetApp.getActive()
        this.ws = this.ss.getSheetByName(SN_EMAIL) || this.ss.getActiveSheet()
    }

    getCharts() {
        const charts = {}
        const [url, gid] = this.ss.getRange(RN_SHEET_URL_OF_CHARTS).getValue().split("#gid=")
        if (!(url && gid)) return charts
        const ss = SpreadsheetApp.openByUrl(url)
        if (!ss) return charts
        const ws = ss.getSheets().find(v => v.getSheetId() == gid)
        if (!ws) return charts
        ws.getCharts().forEach((chart, i) => {
            charts[`{{chart${i + 1}}}`] = chart.getBlob()
        })
        return charts
    }

    sendEmail() {
        const trigger = this.ws.getRange(RN_TRIGGER).getValue()
        if (!trigger) {
            this.ss.toast(`Email trigger is ${trigger}, no email will be sent.`, APP_NAME)
            return
        }
        this.ss.toast(`Sending...`, APP_NAME)
        const { recipient, subject, options } = this.getEmailData()
        try {
            GmailApp.sendEmail(recipient, subject, "", options)
            this.ss.toast(`Email has been sent to ${recipient} successfully.`, APP_NAME, 30)
        } catch (e) {
            this.ss.toast(e.message, APP_NAME, 30)
        }
    }

    createHtmlBody(values, formulas, borders, charts, { width, maxColumnWidth, minColumnWidth, blankRowHeight, columns, bgColor, textColor, borderColor }) {
        let html = `<div style="padding: 12px; background: ${bgColor}; color: ${textColor};"><table style="border-collapse: collapse; width: ${width}px; margin: auto;">`
        values.forEach((row, r) => {
            let tr = `<tr>`
            let style = `padding: 3px 6px; min-width:${minColumnWidth}px; max-width:${maxColumnWidth}px;`
            const singleValueRow = row.filter(v => v !== "")
            const isEmptyRow = row.filter(v => v === "").length === row.length && formulas[r].filter(v => v === "").length === formulas[r].length
            if (isEmptyRow) {
                style = `height: ${blankRowHeight}px;`
                tr += `<td style="${style}" colspan="${row.length}"></td>`
            } else if (singleValueRow.length === 1) {
                const value = singleValueRow[0]
                if (charts[value]) {
                    tr += `<td style="${style}" colspan="${row.length}"><img src="cid:${value}" alt="chart"/></td>`
                } else {
                    tr += `<td style="${style}" colspan="${row.length}">${value}</td>`
                }
            } else {
                row.forEach((cell, c) => {
                    const border = borders[r][c]
                    const formula = formulas[r][c].toLowerCase()
                    const columnPercentage = columns[c]
                    if (columnPercentage) style += `width: ${columnPercentage}%;`
                    if (formula.indexOf("=image(") !== -1) {
                        const image = formulas[r][c].split('("')[1].split('")')[0]
                        cell = `<img src="${image}" alt="image" style="width: 30px;">`
                        style += "text-align: right;"
                    }
                    if (border) {
                        if (border.top) style += `border-top: 1px solid ${borderColor};`
                        if (border.right) style += `border-right: 1px solid ${borderColor};`
                        if (border.bottom) style += `border-bottom: 1px solid ${borderColor};`
                        if (border.left) style += `border-left: 1px solid ${borderColor};`
                        tr += `<td style="${style}padding: 3px 6px;">${cell}</td>`
                    } else {
                        tr += `<td style="${style}">${cell}</td>`
                    }
                })
            }
            tr += "</tr>"

            html += tr
        })
        return html += "</table></div>"
    }

    getEmailData() {
        const recipient = this.ws.getRange(RN_RECIPIENT).getValue()
        const subject = this.ws.getRange(RN_SUBJECT).getValue()
        const cc = this.ws.getRange(RN_CC).getValue()
        const bcc = this.ws.getRange(RN_BCC).getValue()

        const width = this.ws.getRange(RN_WIDTH).getValue()
        const blankRowHeight = this.ws.getRange(RN_BLNAK_ROW_HEIGHT).getValue()
        const minColumnWidth = this.ws.getRange(RN_MIN_COLUMN_WIDTH).getValue()
        const maxColumnWidth = this.ws.getRange(RN_MAX_COLUMN_WIDTH).getValue()
        const columns = this.ws.getRange(RN_COLUMNS).getDisplayValue().split(",").map(v => v.trim())
        const bgColor = this.ws.getRange(RN_BG_COLOR).getValue()
        const textColor = this.ws.getRange(RN_TEXT_COLOR).getValue()
        const borderColor = this.ws.getRange(RN_BORDER_COLOR).getValue()

        const contentRange = this.ws.getRange(RN_CONTENT)
        const contentValues = contentRange.getDisplayValues()
        const contentFormulas = contentRange.getFormulas()

        const contentBorders = []
        for (let r = 0; r < contentValues.length; r++) {
            const value = contentValues[r]
            const contentBorder = []
            for (let c = 0; c < value.length; c++) {
                const border = contentRange.getCell(r + 1, c + 1).getBorder()
                if (!border) {
                    contentBorder.push(border)
                } else {
                    contentBorder.push({
                        top: border.getTop().getBorderStyle(),
                        right: border.getRight().getBorderStyle(),
                        bottom: border.getBottom().getBorderStyle(),
                        left: border.getLeft().getBorderStyle(),
                    })
                }
            }
            contentBorders.push(contentBorder)
        }

        const styles = { width, blankRowHeight, minColumnWidth, maxColumnWidth, columns, bgColor, textColor, borderColor }
        const charts = this.getCharts()
        const htmlBody = this.createHtmlBody(contentValues, contentFormulas, contentBorders, charts, styles)
        const data = {
            recipient,
            subject,
            options: {
                htmlBody,
                cc,
                bcc,
                inlineImages: charts,
            }
        }
        return data
    }
}