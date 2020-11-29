; (function () {
    const PrintUtilityClass = function () {
        return new PrintUtilityClass.init();
    }
    PrintUtilityClass.init = function () {
        $M.init.call(this);
        $D.init.call(this);
        this.template = "";
        this.labelSet = "";
        this.printStatus = "";
        this.process = "";
    }   
    PrintUtilityClass.prototype = {
        createLabelSet_2: function (data)
        {
            var labelSet = new dymo.label.framework.LabelSetBuilder();
            if (data.length > 1) {
                var x = data.length / 2;
                var c = 0;
                x = Math.trunc(x);
                for (var i = 0; i < x; i++) {
                    var record = labelSet.addRecord();
                    record.setText("itemCode1", data[c].ItemCode);
                    record.setText("lotno1", data[c].LotNo);
                    record.setText("qty1", data[c].ReceivedQty);
                    record.setText("dyt1", data[c].ReceivedDate);
                    c++;
                    record.setText("itemCode2", data[c].ItemCode);
                    record.setText("lotno2", data[c].LotNo);
                    record.setText("qty2", data[c].ReceivedQty);
                    record.setText("dyt2", data[c].ReceivedDate);
                    c++;
                }
            } 
            return labelSet;
        },
        createLabelSet_1: function (data) {
            var labelSet = new dymo.label.framework.LabelSetBuilder();
            $.each(data, function (i, v) {
                var record = labelSet.addRecord();
                record.setText("itemCode1", v.ItemCode);
                record.setText("lotno1", v.LotNo);
                record.setText("qty1", v.ReceivedQty);
                record.setText("dyt1", v.ReceivedDate);
            });
            return labelSet;
        },
        createLabelSet_Pahaba: function (data) {
            var labelSet = new dymo.label.framework.LabelSetBuilder();
            $.each(data, function (i, v) {
                var record = labelSet.addRecord();
                record.setText("customer_no", v.ItemCode);
                record.setText("hpp_no", v.ReceivedHPPPNo);
                record.setText("lot_no", v.LotNo);
                record.setText("po_no", v.OrderCode);
                record.setText("description", v.Specification);
                record.setText("qty", v.ReceivedQty);
                record.setText("nw", v.NW);
                record.setText("gw", v.GW);
            });
            return labelSet;
        },
        printLabel: function () {
            let self = this;

            var status = "";
            var labelXml = "";
            let promiseObj = new Promise(function (resolve, reject) {
                $.get("../Label_Templates/" + self.template + ".xml", function (getData, getStatus) {
                    status = getStatus;
                    labelXml = (new XMLSerializer()).serializeToString(getData);
                    if (status == 'success') {
                        try {
                            var label = dymo.label.framework.openLabelXml(labelXml);

                            self.validatePrint(self.process, label).then(function (ok, reject) {
                                if (ok) {
                                    resolve(true);
                                }
                            });
                        }
                        catch (e) {
                            console.log(e.message || e);
                            resolve(false);
                        }
                    } else {
                        self.showWarning("Opening of label template failed!");
                        resolve(false);
                    }
                });
            });
            return promiseObj;
        },
        validatePrint: function (process, label) {
            let self = this;
            let promiseObj = new Promise(function (resolve, reject) {
                $.ajax({
                    type: "POST",
                    url: "/Printer/GetDataPrinters",
                    data: { process: process },
                    dataType: "json",
                    success: function (response) {
                        //console.log(self.responseData);
                        var data = response.data;
                        if (data.status == "success") {
                            var printers = dymo.label.framework.getPrinters();
                            if (printers.length == 0) {
                                self.showWarning("No DYMO printers are installed. Install DYMO printers.");
                                resolve(false);
                            } else {
                                var printerName = "";
                                var isConnected = "";
                                for (var i = 0; i < printers.length; ++i) {
                                    var printer = printers[i];
                                    if (printer.printerType == "LabelWriterPrinter") {
                                        if (printer.name == data.printerName) {
                                            printerName = printer.name;
                                            isConnected = printer.isConnected;
                                            break;
                                        }
                                    }
                                }
                                if (printerName == "") {
                                    self.showWarning("No LabelWriter printers found. Install LabelWriter printer for " + process);
                                    resolve(false);
                                } else {
                                    if (isConnected) {
                                        label.print(printerName, '', self.labelSet);
                                        resolve(true);
                                    } else {
                                        self.showWarning("LabelWriter printer is not connected");
                                        resolve(false);
                                    }
                                }
                            }
                        }
                    }
                });
            });
            
            return promiseObj;
        }
    }
    PrintUtilityClass.init.prototype = $.extend(PrintUtilityClass.prototype, $M.init.prototype, $D.init.prototype);
    PrintUtilityClass.init.prototype = PrintUtilityClass.prototype;
    return window.PrintUtilityClass = window.$PU = PrintUtilityClass;
}());