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

function test() {
    new App().getEmailData()
}

class App {
    constructor() {
        this.ss = SpreadsheetApp.getActive()
        this.ws = this.ss.getSheetByName(SN_EMAIL) || this.ss.getActiveSheet()
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

    createHtmlBody(values, formulas, borders, { width, maxColumnWidth, minColumnWidth, blankRowHeight }) {
        let html = `<div style="padding: 12px;"><table style="border-collapse: collapse; width: ${width}px; margin: auto;">`
        values.forEach((row, r) => {
            let tr = `<tr>`
            let style = `padding: 3px 6px; min-width:${minColumnWidth}px; max-width:${maxColumnWidth}px;`
            const singleValueRow = row.filter(v => v !== "")
            const isEmptyRow = row.filter(v => v === "").length === row.length && formulas[r].filter(v => v === "").length === formulas[r].length
            if (isEmptyRow) {
                style = `height: ${blankRowHeight}px;`
                tr += `<td style="${style}" colspan="${row.length}"></td>`
            } else if (singleValueRow.length === 1) {
                tr += `<td style="${style}" colspan="${row.length}">${singleValueRow[0]}</td>`
            } else {
                row.forEach((cell, c) => {
                    const border = borders[r][c]
                    const formula = formulas[r][c].toLowerCase()
                    if (formula.indexOf("=image(") !== -1) {
                        const image = formulas[r][c].split('("')[1].split('")')[0]
                        cell = `<img src="${image}" alt="image" style="width: 30px;">`
                        style += "text-align: right;"
                    }
                    if (border) {
                        if (border.top) style += "border-top: 1px solid black;"
                        if (border.right) style += "border-right: 1px solid black;"
                        if (border.bottom) style += "border-bottom: 1px solid black;"
                        if (border.left) style += "border-left: 1px solid black;"
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
        const styles = { width, blankRowHeight, minColumnWidth, maxColumnWidth }
        const htmlBody = this.createHtmlBody(contentValues, contentFormulas, contentBorders, styles)
        const data = {
            recipient,
            subject,
            options: {
                htmlBody,
                cc,
                bcc,
            }
        }
        // console.log(data.options.htmlBody)

        return data
    }
}