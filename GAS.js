/**
 * Google Apps Script for Handling Leads and WayForPay Payments
 */

// --- CONFIGURATION ---

function setupEnvironment() {
    const props = PropertiesService.getScriptProperties();
    props.setProperties({
        'MERCHANT_ACCOUNT': 'www_instagram_com2361',
        'MERCHANT_SECRET_KEY': 'a82c3621f0f5ca58a8ffefc594c842ac430080d2',
        'MERCHANT_DOMAIN': 'vova-win.com'
    });
    console.log('Environment variables set successfully!');
}

const CONSTANTS = {
    CURRENCY: 'UAH',
    PRODUCT_NAME: 'Коуч-сесія TRANSFORMER',
    AMOUNT: 2000
};

// --- CORE LOGIC ---

function doGet(e) {
    return handleRequest(e);
}

function doPost(e) {
    return handleRequest(e);
}

function handleRequest(e) {
    const props = PropertiesService.getScriptProperties().getProperties();

    // Parse data at the beginning to use it for sheet selection
    let data;
    if (e.postData && e.postData.contents) {
        try {
            data = JSON.parse(e.postData.contents);
        } catch (err) {
            data = e.parameter;
        }
    } else {
        data = e.parameter;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // Use sheetId from data if provided, otherwise default to 0
    const targetSheetId = data.sheetId ? parseInt(data.sheetId) : 0;
    const sheet = ss.getSheets().find(s => s.getSheetId() === targetSheetId) ||
        ss.getSheets().find(s => s.getSheetId() === 0) ||
        ss.getSheets()[0];

    const callback = e.parameter.callback;

    // 1. Ендпоінт для перевірки статусу (для thanks.html)
    if (e.parameter.action === 'checkStatus' && e.parameter.orderReference) {
        const orderRef = e.parameter.orderReference;
        const rows = sheet.getDataRange().getValues();
        let currentStatus = 'NotFound';

        for (let i = rows.length - 1; i >= 0; i--) {
            if (rows[i][6] == orderRef) { // Column G is orderReference
                currentStatus = rows[i][5]; // Column F is status
                break;
            }
        }

        const result = {
            orderReference: orderRef,
            status: currentStatus,
            isSuccess: currentStatus === 'Оплачено' || currentStatus === 'Зареєстровано'
        };

        if (callback) {
            return ContentService.createTextOutput(callback + '(' + JSON.stringify(result) + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
        }
        return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
    }

    // CASE 1: New Lead from Website (Handles both POST and GET/JSONP)
    if (data.name && data.phone && !data.transactionStatus) {
        const orderReference = data.orderReference || ('ORD-' + new Date().getTime());

        // Handle free or paid leads
        const amount = data.amount !== undefined ? data.amount : (data.isFree ? 0 : CONSTANTS.AMOUNT);
        const status = data.status || (data.isFree ? 'Зареєстровано' : 'Очікує оплати');

        sheet.appendRow([
            data.date || new Date().toLocaleString("uk-UA"),
            data.name,
            data.phone,
            data.telegram || data.contact || '',
            amount,
            data.isFree ? '' : status, // Залишаємо порожнім для безкоштовних, щоб не зсувати колонки
            orderReference,
            data.utm_source || '',
            data.utm_medium || '',
            data.utm_campaign || '',
            data.utm_content || '',
            data.utm_term || ''
        ]);

        const resultData = { status: 'success', orderReference: orderReference };

        if (callback) {
            return ContentService.createTextOutput(callback + '(' + JSON.stringify(resultData) + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
        }
        return ContentService.createTextOutput(JSON.stringify(resultData)).setMimeType(ContentService.MimeType.TEXT);
    }

    // CASE 2: WayForPay Callback
    if (e.postData && e.postData.contents && data.transactionStatus) {
        const wfpData = data;
        const orderRef = wfpData.orderReference;
        const status = wfpData.transactionStatus;

        const rows = sheet.getDataRange().getValues();
        let rowIndex = -1;
        for (let i = 0; i < rows.length; i++) {
            if (rows[i][6] == orderRef) {
                rowIndex = i + 1;
                break;
            }
        }

        if (rowIndex !== -1) {
            if (status === 'Approved') {
                sheet.getRange(rowIndex, 6).setValue('Оплачено');
            } else {
                sheet.getRange(rowIndex, 6).setValue('Відхилено: ' + status);
            }
        }

        const time = Math.floor(new Date().getTime() / 1000);
        const responseSignatureString = [orderRef, 'accept', time].join(';');
        const responseSignature = generateHmacMd5(responseSignatureString, props.MERCHANT_SECRET_KEY);

        const responseBody = {
            orderReference: orderRef,
            status: 'accept',
            time: time,
            signature: responseSignature
        };

        return ContentService.createTextOutput(JSON.stringify(responseBody)).setMimeType(ContentService.MimeType.TEXT);
    }

    return ContentService.createTextOutput(JSON.stringify({ status: 'error' })).setMimeType(ContentService.MimeType.TEXT);
}

function generateHmacMd5(message, key) {
    const signature = Utilities.computeHmacSignature(Utilities.MacAlgorithm.HMAC_MD5, message, key);
    return signature.map(function (chr) {
        return (chr + 256).toString(16).slice(-2);
    }).join('');
}
