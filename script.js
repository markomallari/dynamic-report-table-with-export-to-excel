// Code goes here
var myApp = angular.module("myApp", []);
myApp
  .factory("Excel", function ($window) {
    var uri = "data:application/vnd.ms-excel;base64,",
      template =
        '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40"><head><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{worksheet}</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]--></head><body><table>{table}</table></body></html>',
      base64 = function (s) {
        return $window.btoa(unescape(encodeURIComponent(s)));
      },
      format = function (s, c) {
        return s.replace(/{(\w+)}/g, function (m, p) {
          return c[p];
        });
      };
    return {
      tableToExcel: function (tableId, worksheetName) {
        var table = $(tableId),
          ctx = { worksheet: worksheetName, table: table.html() },
          href = uri + base64(format(template, ctx));
        return href;
      },
    };
  })
  .controller("MyCtrl", function (Excel, $timeout, $scope) {
    $scope.exportToExcel = function (tableId) {
      var exportHref = Excel.tableToExcel(tableId, "WireWorkbenchDataExport");
      $timeout(function () {
        location.href = exportHref;
      }, 100);
    };

    $scope.records = [];

    $scope.convertToTable = function () {
      var textArea = document.getElementById("arrTextArea").value;
      var formattedArray = JSON.parse(JSON.stringify(textArea));
      $scope.records = JSON.parse(formattedArray);
      var copyHeader = JSON.parse(
        JSON.stringify(Object.keys($scope.records[0]))
      );
      $scope.colHeaders1 = copyHeader;
      $scope.colHeaders2 = JSON.parse(JSON.stringify(copyHeader));
    };

    $scope.changeHeader = function (i) {
      console.log($scope.colHeaders1[i]);
      console.log(document.getElementById(`id-${i}`).value);
      $scope.colHeaders1[i] = document.getElementById(`id-${i}`).value;
    };

    $scope.removeHeader = function (i) {
      var arr = JSON.parse(JSON.stringify($scope.colHeaders1));
      var arrRecord = JSON.parse(JSON.stringify($scope.records));
      var val = arr[i];
      console.log(val);
      arr = arr.filter(function (item) {
        return item !== val;
      });

      $scope.colHeaders1 = arr;
      var newRecords = [];
      for (var i = 0; i < arrRecord.length; i++) {
        var details = Object.keys(arrRecord[i])
          .filter((objKey) => objKey !== val)
          .reduce((newObj, key) => {
            newObj[key] = arrRecord[i][key];
            return newObj;
          }, {});
        newRecords.push(details);
      }

      $scope.records = JSON.parse(JSON.stringify(newRecords));
      var copyHeader = JSON.parse(
        JSON.stringify(Object.keys($scope.records[0]))
      );
      $scope.colHeaders1 = JSON.parse(JSON.stringify(copyHeader));
      $scope.colHeaders2 = JSON.parse(JSON.stringify(copyHeader));
    };

    $scope.convert = function (date) {
      if (date) {
        return moment(date).format("MMMM D, Y");
      } else {
        return "";
      }
    };

    $scope.getTotalContacts = function () {
      var total = 0;
      for (var i = 0; i < $scope.records.length; i++) {
        var product = $scope.records[i].total_contacts;
        total += product;
      }
      return total;
    };

    $scope.getTotalValidContacts = function () {
      var total = 0;
      for (var i = 0; i < $scope.records.length; i++) {
        var product = $scope.records[i].valid_contacts;
        total += product;
      }
      return total;
    };
  });
