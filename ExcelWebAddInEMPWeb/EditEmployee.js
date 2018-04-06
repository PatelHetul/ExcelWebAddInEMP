'use strict';

(function () {
	Office.initialize = function (reason) {
		$(document).ready(function () {
			$('#edit').click(BindEmpDetails);
			$('#update').click(SaveEmpDetails);
			$('#opendialog').click(OpenPopup);
			$('#ddlEmployee').change(ClearDocumnet);
		


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


	function BindEmpDetails() {
		Excel.run(function (context) {

			var sheet = context.workbook.worksheets.getActiveWorksheet();
			sheet.getRange("D2").values = "Edit Employee Details";


			sheet.getRange("D2").format.font.set({
				name: "Verdana",
				bold: true,
				size: 18,
				color: "Blue",
			});
			sheet.getRange().format.autofitColumns();

			var e = document.getElementById("ddlEmployee");
			var id = e.options[e.selectedIndex].value;

			if (id == null || id == undefined) {
				return context.sync();
			}

			var result = "hello";
			$.ajax({
				url: '../../api/Employee',
				type: 'GET',
				data: {
					empid: id
				},
				contentType: 'application/json;charset=utf-8'
			}).done(function (data) {
				result = data;
				const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
				const EmployeeTable = currentWorksheet.tables.add("B4:D4", true /*hasHeaders*/);
				EmployeeTable.name = "EmployeeTable";

				EmployeeTable.getHeaderRowRange().values =
					[["  ", " ", "   "]];

				for (var i = 0; i < result.length; i++) {

					EmployeeTable.rows.add(null, [[result[i].employee_Name, ":", result[i].email]]);
				}

				EmployeeTable.getRange().format.font.set({
					name: "Verdana",
					bold: false,
					size: 12,
					color: "Black",
				});
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




	function SaveEmpDetails() {
		Excel.run(function (context) {

			var e = document.getElementById("ddlEmployee");
			var id = e.options[e.selectedIndex].value;


			if (id == null || id == undefined) {
				return context.sync();
			}

			var sourceRange = context.workbook.worksheets.getActiveWorksheet();
			var Empdata= sourceRange.getRange("D5:D10").load("values, rowCount, columnCount");
			
			return context.sync()
				.then(function () {
					var highestValue = Empdata.values;
					var Ename = Empdata.values[0][0];
					var JoingDate = Empdata.values[1][0];
					var Department = Empdata.values[2][0];
					var Email = Empdata.values[3][0];
					var Address = Empdata.values[4][0];
					var Mobile = Empdata.values[5][0];


					var result = "";
					$.ajax({
						url: '../../api/Employee/empid',
						type: 'GET',
						data: {
							empid: id, name: Ename, date: JoingDate, depart: Department, emails: Email, add: Address, mobileno: Mobile
						},
						contentType: 'application/json;charset=utf-8'
					}).done(function (data) {
						result = data;

						if (result == 1) {
							sourceRange.getRange("D12").values = "Employee Update Successfully";
						}
						else {
							sourceRange.getRange("D12").values = "Employee Update Not Successfully";
						}

						sourceRange.getRange("D12").format.font.set({
							name: "Verdana",
							bold: true,
							size: 15,
							color: "RED",
						});
						sourceRange.getRange().format.autofitColumns();

						//	
						return context.sync();
					}).fail(function (status) {
						result = "Could not communicate with the server.";
					});


				})
				.then(context.sync);

			return context.sync();
		}).catch(function (error) {
			console.log("Error: " + error);
			if (error instanceof OfficeExtension.Error) {
				console.log("Debug info: " + JSON.stringify(error.debugInfo));
			}
		});
	}


	function ClearDocumnet() {
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
		}).catch(function (error) {
			console.log("Error: " + error);
			if (error instanceof OfficeExtension.Error) {
				console.log("Debug info: " + JSON.stringify(error.debugInfo));
			}
		});
	}


	//popup  open

	let dialog = null;
	function OpenPopup() {
		debugger;
		Office.context.ui.displayDialogAsync('https://localhost:44369/EditPopup.html',
			{ height: 35, width: 25 },

			function (result) {
				//	debugger;
				dialog = result.value;
				dialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, processMessage);
			});
		//).catch(function (error) {
		//	console.log("Error: " + error);
		//	if (error instanceof OfficeExtension.Error) {
		//		console.log("Debug info: " + JSON.stringify(error.debugInfo));
		//	}
		//});

	}

	function processMessage(arg) {
		$('#user-name').text(arg.message);
		dialog.close();
	}

})();