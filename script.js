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
    $scope.records = [];

    $scope.exportToExcel = function (tableId) {
      var exportHref = Excel.tableToExcel(tableId, "WireWorkbenchDataExport");
      $timeout(function () {
        location.href = exportHref;
      }, 100);
    };

    $scope.convertToTable = function () {
      var textArea = document.getElementById("arrTextArea").value;
      var formattedArray = $scope.cloneItem(textArea);

      try {
        $scope.parseItem(formattedArray);
      } catch (e) {
        $scope.ResponseModel1 = {};
        $scope.ResponseModel1.ResponseAlert1 = true;
        $scope.ResponseModel1.ResponseType1 = "danger";
        $scope.ResponseModel1.ResponseMessage1 =
          "Invalid Array Object, Please use JSON formatter to correct white spaces and characters, etc";

        $timeout(function () {
          $scope.ResponseModel1.ResponseAlert1 = false;
        }, 5000);
        return;
      }
      //for body row datas
      $scope.records = $scope.parseItem(formattedArray);

      //for header row datas
      const formatted = $scope.parseItem(formattedArray);
      let uniqueHeader = [];
      formatted.map((val) => {
        uniqueHeader = [...uniqueHeader, ...Object.keys(val)];
      });
      const headers = uniqueHeader.filter(
        (item, index) => uniqueHeader.indexOf(item) === index
      );

      var copyHeader = $scope.cloneItem(headers);
      $scope.colHeaders1 = copyHeader;
      $scope.colHeaders2 = $scope.cloneItem(copyHeader);
    };

    $scope.changeHeader = function (i) {
      $scope.ResponseModel = {};
      var newDom = document.getElementById(`id-${i}`);
      if ($scope.findDuplicateHeader($scope.colHeaders1, newDom.value)) {
        newDom.classList.add("error");
        $scope.ResponseModel.ResponseAlert = true;
        $scope.ResponseModel.ResponseType = "danger";
        $scope.ResponseModel.ResponseMessage =
          "Same header name is not permitted";
      } else {
        newDom.classList.remove("error");
        $scope.ResponseModel.ResponseAlert = false;
        $scope.colHeaders1[i] = newDom.value;
      }
    };

    $scope.findDuplicateHeader = function (arr, val) {
      return arr.includes(val);
    };

    $scope.removeHeader = function (i) {
      /** removing key on colHeader1 **/
      var arr = $scope.cloneItem($scope.colHeaders1);
      var val = arr[i];
      arr = arr.filter(function (item) {
        return item !== val;
      });
      $scope.colHeaders1 = arr;
      /** removing key on colHeader2 **/
      var arr2 = $scope.cloneItem($scope.colHeaders2);
      arr2 = arr2.filter(function (item, key) {
        return key !== i;
      });
      $scope.colHeaders2 = arr2;
      /** updating row records **/
      var arrRecord = $scope.cloneItem($scope.records);
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
      $scope.records = $scope.cloneItem(newRecords);
    };

    $scope.dataChecker = function (val) {
      if ($scope.isDateValid(val)) {
        return moment(val).format("MMMM D, Y");
      } else {
        return val;
      }
    };

    $scope.isDateValid = function (val) {
      var formats = [moment.ISO_8601, "MM/DD/YYYY  :)  HH*mm*ss"];
      return moment(val, formats, true).isValid(); // true
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

    $scope.cloneItem = function (item) {
      return $scope.parseItem(JSON.stringify(item));
    };

    $scope.parseItem = function (item) {
      return JSON.parse(item);
    };
  });
