// test
// const token = "2126091524:AAEFX8-yMwiGkbz-jMsRgu8dks0H4b5_i2g";
// const chat_id = "-4129828850";
// live
// const token = "7060991688:AAHBP3MpFYwt4FXtJH-ysJmd-M2e4Wa2wZQ";
// const chat_id = "-4184337807"
// local
const token = "6929001417:AAGd-7cNM4WrLzMsrDLpgQKtUnNioQUKiNE";
const chat_id = "1827717290";

const url = "https://api.telegram.org/bot" + token + "/sendMessage";
function sendMessageTelegram(message) {
    var payload = {
        "chat_id": chat_id,
        "text": message,
        "parse_mode": "HTML",
        "disable_web_page_preview": true
    };
    var options = {
        "method": "post",
        "payload": payload
    };
    UrlFetchApp.fetch(url, options);
}
function getYesterday() {
    var yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);
    // Lấy ngày và tháng của ngày hôm qua
    var dayChoose = yesterday.getDate();
    var monthChoose = yesterday.getMonth() + 1; // Lấy tháng cần +1 vì hàm getMonth() trả về từ 0-11
    var yearChoose = yesterday.getFullYear();
    // Định dạng ngày và tháng thành chuỗi với định dạng dd/mm
    var formattedDate = (dayChoose < 10 ? '0' : '') + dayChoose + '/' + (monthChoose < 10 ? '0' : '') + monthChoose + '/' + yearChoose;
    //var formattedDate = "01/02/2024";

    return formattedDate;
}
// Hàm để định dạng ngày thành "dd/MM/yyyy"
function formatDate(date) {
    var day = date.getDate();
    var month = date.getMonth() + 1; // Lưu ý: Tháng bắt đầu từ 0
    var year = date.getFullYear();
    return (day < 10 ? '0' : '') + day + '/' + (month < 10 ? '0' : '') + month + '/' + year;
}
function findColumnWithYesterdayDate(rows) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("5. Monthly version 3.1_231222");
    var yesterday = new Date();
    formattedDate = yesterday.setDate(yesterday.getDate() - 1);
    var lastColumn = sheet.getLastColumn();

    // Lặp qua các cột để tìm cột chứa giá trị ở hàng có ngày là ngày hôm qua
    for (var column = 1; column <= lastColumn; column++) {
        var value_start = sheet.getRange(rows + 1, column).getValue(); // Lấy giá trị từ ô tại hàng và cột đang xét
        var value_end = sheet.getRange(rows + 2, column).getValue(); // Lấy giá trị từ ô tại hàng và cột đang xét
        // Chuyển đổi giá trị thành đối tượng ngày
        var start_at = new Date(value_start);
        var end_at = new Date(value_end);

        // Kiểm tra nếu formattedDate nằm trong khoảng từ start_at đến end_at
        if (formattedDate >= start_at && formattedDate <= end_at) {
            // Nếu formattedDate nằm trong khoảng, trả về số cột đang xét
            return column;
        }
    }
    return -1;
}

function getInfoColumn() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("5. Monthly version 3.1_231222");
    if (!sheet) {
        Logger.log("Không tìm thấy sheet có tên '5. Monthly version 3.1_231222'");
        return;
    }
    var arrayName = ["Total","Anhnm","SơnPV","TràNH","NgocNA"];
    var cotA = sheet.getRange('A1:A').getValues(); // Lấy giá trị của cột A từ hàng 1 đến hết
    var finalResult = [];
    // Lặp qua từng tên trong mảng arrayName
    for (var i = 0; i < arrayName.length; i++) {
        var currentName = arrayName[i]; // Lấy tên hiện tại trong mảng arrayName
        var rowIndex = cotA.findIndex(function(row) {
            return row[0] === currentName; // Tìm vị trí của tên trong cột A
        });
        if (rowIndex !== -1) {
            // Nếu tìm thấy tên trong cột A, thêm cặp tên và vị trí hàng vào mảng kết quả
            finalResult.push({ [currentName]: rowIndex + 1 });
        } else {
            // Nếu không tìm thấy, thông báo rằng không tìm thấy
            Logger.log("Tên: " + currentName + " không tìm thấy trong cột A");
        }
    }
    // Tạo một mảng mới chứa các cặp tên và giá trị
    var newResult = finalResult.map(function(item) {
        var key = Object.keys(item)[0]; // Lấy tên từ khóa
        var value = item[key]; // Lấy giá trị từ khóa
        var result = {}; // Tạo một đối tượng mới để chứa cặp tên và giá trị
        result[key] = {}; // Khởi tạo đối tượng con cho từng tên
        var columnYesterday = findColumnWithYesterdayDate(value);

        // Lặp qua 27 hàng kế tiếp và lấy giá trị từ cột E và B
        if(columnYesterday != -1){
            for (var j = value; j < value + 27; j++) {
                if(j > value + 2){
                    var rowFound = sheet.getRange(j, columnYesterday, 1, 1).getValue(); // Lấy giá trị từ cột E
                    var columnB = sheet.getRange(j, 2, 1, 1).getValue(); // Lấy giá trị từ cột B
                    result[key][columnB] = rowFound; // Gán giá trị vào đối tượng kết quả
                }
            }
        }

        return result; // Trả về đối tượng mới chứa cặp tên và giá trị
    });
    return newResult;
}
// Hàm định dạng số với dấu phẩy
function formatNumber(number) {
    return number.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}
// Hàm chuyển đổi số thành phần trăm
function toPercentage(number) {
    return (number * 100).toFixed(2) + '%';
}
function sendDataToTelegram(){
    var times = getYesterday();
    var result = getInfoColumn();
    var message = `<b>Báo cáo CNHTD ngày ${times}</b>`;
    // Biến đếm số thứ tự của agent
    var count = 1;
    // Thêm thông tin của mỗi đối tượng từ mảng result vào message
    result.forEach(function(agent){
        var agentName = Object.keys(agent)[0];
        var agentInfo = agent[agentName];
        message += `
        <b>\n  ${count}. ${agentName}: </b>
      Thông tin chi tiết:`
        message +=
            (agentInfo.hasOwnProperty('RE') ? `\n       - <b>RE:</b> ${formatNumber(agentInfo.RE)}\n` : '') +
            (agentInfo.hasOwnProperty('L0') ? `       - <b>L0:</b> ${formatNumber(agentInfo.L0)}\n` : '') +
            (agentInfo.hasOwnProperty('L2') ? `       - <b>L2:</b> ${formatNumber(agentInfo.L2)}\n` : '') +
            (agentInfo.hasOwnProperty('L4A ') ? `       - <b>L4A:</b> ${formatNumber(agentInfo['L4A '])}\n` : '') +
            (agentInfo.hasOwnProperty('L8') ? `       - <b>L8:</b> ${formatNumber(agentInfo.L8)}\n` : '') +
            (agentInfo.hasOwnProperty('L9') ? `       - <b>L9:</b> ${formatNumber(agentInfo.L9)}\n` : '') +
            (agentInfo.hasOwnProperty('k') ? `       - K: ${formatNumber(agentInfo.k)}\n` : '')
        message += `
       Tỷ lệ:
        - <b>L9/L8</b>: ${agentInfo.hasOwnProperty('L9/L8') ? toPercentage(agentInfo['L9/L8']) : 'N/A'}
        - <b>L8/L1:</b> ${agentInfo.hasOwnProperty('L8/L1') ? toPercentage(agentInfo['L8/L1']) : 'N/A'}
        - <b>L4A/L1:</b> ${agentInfo.hasOwnProperty('L4A/L1') ? toPercentage(agentInfo['L4A/L1']) : 'N/A'}
        - <b>L2/L1:</b> ${agentInfo.hasOwnProperty('L2/L1') ? toPercentage(agentInfo['L2/L1']) : 'N/A'}`;
        // Tăng biến đếm số thứ tự của agent
        count++;
    });
    Logger.log(message);
    sendMessageTelegram(message);

}
// biểu đồ
function createChartImageAndUploadToDrive() {
    var spreadsheetId = "1SJfJ7bDb2Sc11c8vSTRv5crSoFYce7qXKtPypWO9xrk";
    var sheetName = "DASH"; // tên sheet chứa dữ liệu biểu đồ

    // Lấy dữ liệu từ Google Sheets
    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    var sheet = spreadsheet.getSheetByName(sheetName);
    var charts = sheet.getCharts();

    for (var i = 0; i < charts.length; i++) {
        let chartName =  charts[i].getOptions().get('title');
        var chart = getChartByName(sheet, chartName);
        if (!chart) {
            Logger.log("Không tìm thấy biểu đồ có tên '" + chartName + "' trên sheet '" + sheetName + "'.");
            return;
        }
        // Tạo hình ảnh của biểu đồ
        var chartImage = chart.getAs('image/png');
        // Tạo tên file và tải hình ảnh lên Google Drive
        var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
        var fileName = chartName + "_" + timestamp + ".png";
        var folder = DriveApp.getFolderById("11OwuBzZr4BKgc5PLBNl5g7ZRr_KlNzrL"); // Thay YOUR_DRIVE_FOLDER_ID bằng ID thư mục bạn muốn lưu trữ hình ảnh
        var file = folder.createFile(chartImage.setName(fileName));
        // Lấy URL của hình ảnh trên Google Drive
        var fileUrl = file.getUrl();
        // Gửi URL hình ảnh đến Telegram
        sendMessageTelegramImage(fileUrl);
    }

    // Gọi hàm callback sau khi hoàn thành tất cả công việc
    callback();

}

// Hàm lấy biểu đồ theo tên
function getChartByName(sheet, chartName) {
    var charts = sheet.getCharts();
    for (var i = 0; i < charts.length; i++) {
        if (charts[i].getOptions().get('title') === chartName) {
            return charts[i];
        }
    }
    return null;
}
const url_image = "https://api.telegram.org/bot" + token + "/sendPhoto";
function extractFileIdFromDriveURL(driveURL) {
    // Định dạng biểu thức chính quy để tìm kiếm ID tệp từ URL Google Drive
    var regex = /\/file\/d\/([a-zA-Z0-9-_]+)\//;

    // Sử dụng biểu thức chính quy để tìm ID tệp
    var match = regex.exec(driveURL);

    // Nếu có sự khớp, trả về ID tệp
    if (match && match[1]) {
        return match[1];
    } else {
        // Nếu không tìm thấy ID, trả về null hoặc thông báo lỗi
        return null;
    }
}
// Hàm tải ảnh từ Google Drive
function getImageFromDrive(imageUrl) {
    var fileId = extractFileIdFromDriveURL(imageUrl);
    //var fileId = "1JwAlINS_7jrYRFu2M7Rybv724KNok_S-";
    var downloadUrl = DriveApp.getFileById(fileId).getBlob();
    //var downloadUrl = "https://drive.google.com/uc?export=view&amp;&id=" + fileId;
    return downloadUrl;
}

function sendMessageTelegramImage(imageUrl) {
    var imageBlob = getImageFromDrive(imageUrl);
    //var imageUrl = "https://drive.google.com/uc?export=view&id=19AHoVAXGH7avd5G2SG3jAN_N1ybSEl4c";
    try {
        var payload = {
            "chat_id": chat_id,
            "photo": imageBlob
        };
        var options = {
            "method": "post",
            "payload": payload,
            "muteHttpExceptions": true
        };
        var response =UrlFetchApp.fetch(url_image, options);
        Logger.log(response.getContentText());
    } catch (e) {
        Logger.log("Lỗi: " + e);
    }
}

function sendDataDaily() {
    // Gọi trước hàm sendDataToTelegram()
    sendDataToTelegram();
    // Gọi sau hàm sendDataToTelegram(), sử dụng hàm callback
    createChartImageAndUploadToDrive(function() {
        Logger.log("createChartImageAndUploadToDrive() đã hoàn thành!");
    });
}









