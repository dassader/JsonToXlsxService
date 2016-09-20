var parse = (function () {
    return function (tableDom, name) {
        tableDom = $(tableDom);

        var sheet = {
            name: name,
            rows: []
        };


        tableDom.find("tr").each(function (index, targetTr) {
            targetTr = $(targetTr);

            if(!targetTr.is(":visible")) {
                return;
            }

            var row = {
                cells: []
            };

            targetTr.find("th, td").each(function (index, targetTd) {
                targetTd = $(targetTd);

                if(!targetTd.is(":visible")) {
                    return;
                }

                row.cells.push({
                    colspan: targetTd.attr("colspan") || 0,
                    value: targetTd.text().trim() + inputsToText(targetTd.find("input"))
                });
            });
            sheet.rows.push(row);
        });

        function inputsToText(inputs) {
            var value = "";

            inputs.each(function (index, target) {
                value += $(target).val() + " ";
            });

            return value;
        }

        return sheet;
    }


})();