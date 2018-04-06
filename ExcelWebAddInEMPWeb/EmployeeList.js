'use strict';

(function () {
	Office.initialize = function (reason) {
		$(document).ready(function () {
			$('#msg').click(PrintMsg);
			$('#emplist').click(Displayemplist);
			$('#empchart').click(DisplayempChart);
			
			Excel.run(function (context) {

				var sheet = context.workbook.worksheets.getActiveWorksheet();
				sheet.getRange("A1:H20").values = "";
				sheet.getRange().format.font.set({
					name: "Calibri",
					bold: false,
					size: 11,
					color: "Black",
				});

				return context.sync();
			})
		});
	};

	function PrintMsg() {
		Excel.run(function (context) {

			var sheet = context.workbook.worksheets.getActiveWorksheet();
			sheet.getRange("C2").values = "Employees List";
			
			sheet.getRange("C2").format.font.set({
				name: "Verdana",
				bold: true,
				size: 18,
				color: "Blue",
			});
			sheet.getRange().format.autofitColumns();
			return context.sync();
		}).catch(function (error) {
			console.log("Error: " + error);
			if (error instanceof OfficeExtension.Error) {
				console.log("Debug info: " + JSON.stringify(error.debugInfo));
			}
		});
	}

	function Displayemplist() {
		Excel.run(function (context) {

			var result = "hello";
			$.ajax({
				url: '../../api/Employee',
				type: 'GET',
				data: {
					empid: 0
				},
				contentType: 'application/json;charset=utf-8'
			}).done(function (data) {
				result = data;
				const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
				const EmployeeTable = currentWorksheet.tables.add("A4:G4", true /*hasHeaders*/);
				EmployeeTable.name = "EmployeeTable";

				EmployeeTable.getHeaderRowRange().values =
					[["Employee Name", "Department", "Joining Date", "Address","Email","Mobile No","Salary"]];
				
				for (var i = 0; i < result.length; i++) {
				
					EmployeeTable.rows.add(null, [[result[i].employee_Name, result[i].department_Name, result[i].joiningDate, result[i].address, result[i].email, result[i].mobileNo, result[i].Salary]]);
				}
				

				
				EmployeeTable.getRange().format.font.set({
					name: "Verdana",
					bold: false,
					size: 12,
					color: "Black",
				});
				EmployeeTable.columns.getItemAt(6).getRange().numberFormat = [['€#,##0.00']];
				EmployeeTable.getRange().format.autofitColumns();
				EmployeeTable.getRange().format.autofitRows();

				return context.sync();
			}).fail(function (status) {
				result = "Could not communicate with the server.";
			});

			return context.sync();
		}).catch(function (error) {
			console.log("Error: " + error);
			if (error instanceof OfficeExtension.Error) {
				console.log("Debug info: " + JSON.stringify(error.debugInfo));
			}
		});
	}

	function DisplayempChart() {
		Excel.run(function (context) {

			const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
			const EmployeeTable = currentWorksheet.tables.getItem('EmployeeTable');
			const dataRange = EmployeeTable.getDataBodyRange();

			let chart = currentWorksheet.charts.add('ColumnClustered', dataRange, 'auto');

			chart.setPosition("A15", "F45");
			chart.title.text = "Expenses";
			chart.legend.position = "right"
			chart.legend.format.fill.setSolidColor("white");
			chart.dataLabels.format.font.size = 12;
			chart.dataLabels.format.font.color = "black";
			chart.series.getItemAt(0).name = 'Value in €';

			return context.sync();
		}).catch(function (error) {
			console.log("Error: " + error);
			if (error instanceof OfficeExtension.Error) {
				console.log("Debug info: " + JSON.stringify(error.debugInfo));
			}
		});
	}
})();