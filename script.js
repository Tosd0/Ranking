document.addEventListener('DOMContentLoaded', () => {
    const csvFileInput = document.getElementById('csvFileInput');
    const processButton = document.getElementById('processButton');
    const messageDiv = document.getElementById('message');

    csvFileInput.addEventListener('change', () => {
        if (csvFileInput.files.length > 0) {
            processButton.disabled = false; // 允许点击处理按钮
            messageDiv.textContent = ""; // 清空消息
        } else {
            processButton.disabled = true; // 禁用处理按钮
        }
    });

    processButton.addEventListener('click', () => {
        const file = csvFileInput.files[0];
        if (!file) {
            messageDiv.textContent = "请先选择 CSV 文件。";
            return;
        }

        messageDiv.textContent = "正在处理 CSV 文件...";
        processButton.disabled = true; // 处理时禁用按钮

        const reader = new FileReader();
        reader.onload = function(event) {
            const csvText = event.target.result;

            Papa.parse(csvText, {
                header: false, // 假设 CSV 没有 header，或者不需要作为 header
                skipEmptyLines: true,
                complete: function(results) {
                    const csvData = results.data;
                    if (csvData && csvData.length > 0) {
                        try {
                            const schoolData = processCSV(csvData);
                            const rankedSchools = rankSchools(schoolData);
                            outputRankingToExcel(rankedSchools);
                            messageDiv.textContent = "Excel 文件生成成功，已开始下载。";
                        } catch (error) {
                            console.error("处理 CSV 数据出错:", error);
                            messageDiv.textContent = "处理 CSV 数据出错，请查看控制台错误信息。";
                        }
                    } else {
                        messageDiv.textContent = "CSV 文件内容为空或解析失败。";
                    }
                    processButton.disabled = false; // 处理完成后重新启用按钮
                },
                error: function(error) {
                    console.error("CSV 解析错误:", error);
                    messageDiv.textContent = "CSV 文件解析错误，请检查文件格式。";
                    processButton.disabled = false; // 出错后重新启用按钮
                }
            });
        };
        reader.onerror = function() {
            console.error("文件读取错误");
            messageDiv.textContent = "文件读取错误，请重试。";
            processButton.disabled = false; // 出错后重新启用按钮
        };
        reader.readAsText(file, 'utf-8'); // 以 UTF-8 编码读取文件
    });

    function calculateScores(dataRow) {
        // *** JavaScript 版本的 calculate_scores 函数 ***
        const school_b = dataRow[1];
        const school_c = dataRow[2];
        const j_value = dataRow[7];
        const k = parseInt(dataRow[8]) || 0;
        const l = parseInt(dataRow[9]) || 0;
        const m = parseInt(dataRow[10]) || 0;
        const n_val = parseInt(dataRow[11]) || 0;
        const o_val = parseInt(dataRow[12]) || 0;
        const p = parseInt(dataRow[13]) || 0;

        let regulation_score_b, survival_score_b, regulation_score_c, survival_score_c;

        if (j_value === 'Yes') {
            regulation_score_b = k;
            survival_score_b = n_val;
            regulation_score_c = o_val;
            survival_score_c = l;
        } else if (j_value === 'No') {
            regulation_score_b = n_val;
            survival_score_b = k;
            regulation_score_c = l;
            survival_score_c = o_val;
        } else {
            throw new Error(`J列的值 '${j_value}' 既不是 'Yes' 也不是 'No'，请检查数据。`);
        }

        const small_score_b = regulation_score_b + survival_score_b;
        const small_score_c = regulation_score_c + survival_score_c;

        let winner = '';
        if (small_score_b > small_score_c) {
            winner = 'B';
        } else if (small_score_c > small_score_b) {
            winner = 'C';
        } else { // 小分相等时
            if (j_value === 'Yes') {
                if (k === 5) {
                    if (m < p) {
                        winner = 'B';
                    } else if (m > p) {
                        winner = 'C';
                    }
                } else { // k != 5
                    if (m < p) {
                        winner = 'C';
                    } else if (m > p) {
                        winner = 'B';
                    }
                }
            } else if (j_value === 'No') {
                if (k === 5) {
                    if (m < p) {
                        winner = 'C';
                    } else if (m > p) {
                        winner = 'B';
                    }
                } else { // k != 5
                    if (m < p) {
                        winner = 'B';
                    } else if (m > p) {
                        winner = 'C';
                    }
                }
            } else { // 理论上不应该到这里
                winner = 'Tie';
            }
        }

        return {
            school_b: school_b,
            school_c: school_c,
            j_value: j_value,
            regulation_score_b: regulation_score_b,
            survival_score_b: survival_score_b,
            regulation_score_c: regulation_score_c,
            survival_score_c: survival_score_c,
            small_score_b: small_score_b,
            small_score_c: small_score_c,
            winner: winner
        };
    }

    function processCSV(csvData) {
        // *** JavaScript 版本的 process_csv 函数 ***
        const schoolData = {};
        const rowResults = [];

        for (let i = 1; i < csvData.length; i++) { // 从第二行开始，跳过可能的 header
            const row = csvData[i];
            if (row.length < 16) {
                console.warn(`警告: CSV 文件行数据列数不足，跳过该行: ${row}`);
                continue;
            }

            try {
                const result = calculateScores(row);
                rowResults.push(result);

                const school_b = result.school_b;
                const school_c = result.school_c;
                const winner = result.winner;
                const regulation_score_b = result.regulation_score_b;
                const survival_score_b = result.survival_score_b;
                const regulation_score_c = result.regulation_score_c;
                const survival_score_c = result.survival_score_c;

                // 初始化学校数据
                if (!schoolData[school_b]) {
                    schoolData[school_b] = {
                        '积分': 0,
                        '小分总和': 0,
                        '求生总得分': 0,
                        '监管总得分': 0,
                        '出现次数': 0
                    };
                }
                // 进行累加，并在累加前检查是否为 NaN
                if (!isNaN(result.small_score_b)) {
                    schoolData[school_b]['小分总和'] += result.small_score_b;
                }
                if (!isNaN(survival_score_b)) {
                    schoolData[school_b]['求生总得分'] += survival_score_b;
                }
                if (!isNaN(regulation_score_b)) {
                    schoolData[school_b]['监管总得分'] += regulation_score_b;
                }
                schoolData[school_b]['出现次数'] += 1;
                if (winner === 'B') {
                    schoolData[school_b]['积分'] += 3;
                }

                // 对 school_c 做同样的处理
                if (!schoolData[school_c]) {
                    schoolData[school_c] = {
                        '积分': 0,
                        '小分总和': 0,
                        '求生总得分': 0,
                        '监管总得分': 0,
                        '出现次数': 0
                    };
                }
                if (!isNaN(result.small_score_c)) {
                    schoolData[school_c]['小分总和'] += result.small_score_c;
                }
                if (!isNaN(survival_score_c)) {
                    schoolData[school_c]['求生总得分'] += survival_score_c;
                }
                if (!isNaN(regulation_score_c)) {
                    schoolData[school_c]['监管总得分'] += regulation_score_c;
                }
                schoolData[school_c]['出现次数'] += 1;
                if (winner === 'C') {
                    schoolData[school_c]['积分'] += 3;
                }

            } catch (e) {
                console.error(`处理行时出错 (学校 B: ${row[1]}, 学校 C: ${row[2]}): ${e}`);
            }
        }

        // 计算平均得分
        for (const schoolName in schoolData) {
            const data = schoolData[schoolName];
            const appearances = data['出现次数'];
            if (appearances > 0) {
                data['求生局均得分'] = (data['求生总得分'] / appearances).toFixed(3);
                data['监管局均得分'] = (data['监管总得分'] / appearances).toFixed(3);
            } else {
                data['求生局均得分'] = 0; // 出现次数为 0 时，均分设为 0
                data['监管局均得分'] = 0;
            }
        }

        return schoolData;
    }

    function rankSchools(schoolData) {
        const rankedSchools = [];
        for (const schoolName in schoolData) {
            const data = schoolData[schoolName];
            rankedSchools.push({
                '排名': 0, // 初始排名为 0
                '学校/队伍名称': schoolName,
                '积分': data['积分'],
                '小分': data['小分总和'],
                '求生局均得分': data['求生局均得分'],
                '监管局均得分': data['监管局均得分']
            });
        }

        rankedSchools.sort((a, b) => {
            if (b['积分'] !== a['积分']) {
                return b['积分'] - a['积分']; // 积分降序
            }
            if (b['小分'] !== a['小分']) {
                return b['小分'] - a['小分'];     // 小分降序
            }
            if (b['求生局均得分'] !== a['求生局均得分']) {
                return b['求生局均得分'] - a['求生局均得分']; // 求生局均得分降序
            }
            return b['监管局均得分'] - a['监管局均得分']; // 监管局均得分降序
        });

        let rank = 1;
        for (let i = 0; i < rankedSchools.length; i++) {
            if (i > 0 && (rankedSchools[i]['积分'] !== rankedSchools[i - 1]['积分'] ||
                rankedSchools[i]['小分'] !== rankedSchools[i - 1]['小分'] ||
                rankedSchools[i]['求生局均得分'] !== rankedSchools[i - 1]['求生局均得分'] ||
                rankedSchools[i]['监管局均得分'] !== rankedSchools[i - 1]['监管局均得分'])) {
                rank = i + 1;
            }
            rankedSchools[i]['排名'] = rank;
        }

        return rankedSchools;
    }

    function outputRankingToExcel(rankedSchools) {
        // *** JavaScript 版本的 output_ranking_to_excel 函数 ***
        const worksheetData = [
            ["排名", "学校/队伍名称", "积分", "小分", "求生局均得分", "监管局均得分"], // Header row
            ...rankedSchools.map(school => [ // 数据行
                school['排名'],
                school['学校/队伍名称'],
                school['积分'],
                school['小分'],
                parseFloat(school['求生局均得分']),
                parseFloat(school['监管局均得分'])
            ])
        ];

        const footerLine1 = "排名程序编写：TO";
        const footerLine2 = "若有数据问题请联系当场裁判";

        const footerRow1 = new Array(worksheetData[0].length).fill(null);
        footerRow1[0] = footerLine1;
        worksheetData.push(footerRow1);

        const footerRow2 = new Array(worksheetData[0].length).fill(null);
        footerRow2[0] = footerLine2;
        worksheetData.push(footerRow2);

        const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "排名表");

        // 合并单元格 - 第一行footer
        let mergeCell1 = {
            s: { r: worksheetData.length - 2, c: 0 }, // 倒数第二行，开始单元格
            e: { r: worksheetData.length - 2, c: worksheetData[0].length - 1 } // 倒数第二行，结束单元格
        };
        if (!worksheet['!merges']) worksheet['!merges'] = [];
        worksheet['!merges'].push(mergeCell1);

        // 样式 - 第一行footer：居中对齐和加粗
        let cellAddress1 = XLSX.utils.encode_cell({ r: worksheetData.length - 2, c: 0 });
        if (!worksheet[cellAddress1]) worksheet[cellAddress1] = {};
        if (!worksheet[cellAddress1].s) worksheet[cellAddress1].s = {};
        worksheet[cellAddress1].s.alignment = { horizontal: "center", vertical: "center" };
        worksheet[cellAddress1].s.font = { bold: true };

        // 合并单元格 - 第二行footer
        let mergeCell2 = {
            s: { r: worksheetData.length - 1, c: 0 }, // 最后一行，开始单元格
            e: { r: worksheetData.length - 1, c: worksheetData[0].length - 1 } // 最后一行，结束单元格
        };
        worksheet['!merges'].push(mergeCell2);

        // 样式 - 第二行footer：居中对齐和红色
        let cellAddress2 = XLSX.utils.encode_cell({ r: worksheetData.length - 1, c: 0 });
        if (!worksheet[cellAddress2]) worksheet[cellAddress2] = {};
        if (!worksheet[cellAddress2].s) worksheet[cellAddress2].s = {};
        worksheet[cellAddress2].s.alignment = { horizontal: "center", vertical: "center" };
        worksheet[cellAddress2].s.font = { color: { rgb: "FFFF0000" } }; // 红色 RGB 值

        const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'ranking_table.xlsx';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }
});