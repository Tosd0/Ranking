document.addEventListener('DOMContentLoaded', () => {
    const csvFileInput = document.getElementById('csvFileInput');
    const processButton = document.getElementById('processButton'); // 原按钮
    const groupedButton = document.getElementById('groupedButton'); // 新按钮
    const messageDiv = document.getElementById('message');

    let csvDataCache = null; // 用于存储解析后的原始 CSV 数据，以便两个按钮都能使用

    // 当选择了文件时，启用按钮
    csvFileInput.addEventListener('change', () => {
        if (csvFileInput.files.length > 0) {
            processButton.disabled = false;
            groupedButton.disabled = false;
            messageDiv.textContent = "";
        } else {
            processButton.disabled = true;
            groupedButton.disabled = true;
        }
    });

    // ======= 原来的逻辑，生成并下载「积分榜」 =======
    processButton.addEventListener('click', () => {
        const file = csvFileInput.files[0];
        if (!file) {
            messageDiv.textContent = "请先选择 CSV 文件。";
            return;
        }

        messageDiv.textContent = "正在处理 CSV 文件（原积分榜逻辑）...";
        processButton.disabled = true;
        groupedButton.disabled = true;

        const reader = new FileReader();
        reader.onload = function(event) {
            const csvText = event.target.result;

            Papa.parse(csvText, {
                header: false,
                skipEmptyLines: true,
                complete: function(results) {
                    csvDataCache = results.data; // 缓存一下
                    if (csvDataCache && csvDataCache.length > 0) {
                        try {
                            const { schoolData, roundData } = processCSV(csvDataCache);
                            const rankedSchools = rankSchools(schoolData);
                            outputRankingToExcel(rankedSchools, roundData); // 原有的输出函数
                            messageDiv.textContent = "Excel 文件生成成功（原积分表），已开始下载。";
                        } catch (error) {
                            console.error("处理 CSV 数据出错:", error);
                            messageDiv.textContent = "处理 CSV 数据出错，请查看控制台错误信息。";
                        }
                    } else {
                        messageDiv.textContent = "CSV 文件内容为空或解析失败。";
                    }
                    processButton.disabled = false;
                    groupedButton.disabled = false;
                },
                error: function(error) {
                    console.error("CSV 解析错误:", error);
                    messageDiv.textContent = "CSV 文件解析错误，请检查文件格式。";
                    processButton.disabled = false;
                    groupedButton.disabled = false;
                }
            });
        };
        reader.onerror = function() {
            console.error("文件读取错误");
            messageDiv.textContent = "文件读取错误，请重试。";
            processButton.disabled = false;
            groupedButton.disabled = false;
        };
        reader.readAsText(file, 'utf-8');
    });

    // ======= 新按钮逻辑，生成并下载「已分组表格」 =======
    groupedButton.addEventListener('click', () => {
        const file = csvFileInput.files[0];
        if (!file) {
            messageDiv.textContent = "请先选择 CSV 文件。";
            return;
        }

        messageDiv.textContent = "正在处理 CSV 文件（分组表格逻辑）...";
        processButton.disabled = true;
        groupedButton.disabled = true;

        const reader = new FileReader();
        reader.onload = function(event) {
            const csvText = event.target.result;

            Papa.parse(csvText, {
                header: false,
                skipEmptyLines: true,
                complete: function(results) {
                    csvDataCache = results.data; // 缓存一下
                    if (csvDataCache && csvDataCache.length > 0) {
                        try {
                            // 做跟上面一样的计算
                            const { schoolData, roundData } = processCSV(csvDataCache);
                            const rankedSchools = rankSchools(schoolData);
                            // 调用新的「分组表格」输出函数（你贴出的那一段）
                            outputGroupedRankingToExcel(rankedSchools, roundData);
                            messageDiv.textContent = "Excel 文件生成成功（已分组表格），已开始下载。";
                        } catch (error) {
                            console.error("处理 CSV 数据出错:", error);
                            messageDiv.textContent = "处理 CSV 数据出错，请查看控制台错误信息。";
                        }
                    } else {
                        messageDiv.textContent = "CSV 文件内容为空或解析失败。";
                    }
                    processButton.disabled = false;
                    groupedButton.disabled = false;
                },
                error: function(error) {
                    console.error("CSV 解析错误:", error);
                    messageDiv.textContent = "CSV 文件解析错误，请检查文件格式。";
                    processButton.disabled = false;
                    groupedButton.disabled = false;
                }
            });
        };
        reader.onerror = function() {
            console.error("文件读取错误");
            messageDiv.textContent = "文件读取错误，请重试。";
            processButton.disabled = false;
            groupedButton.disabled = false;
        };
        reader.readAsText(file, 'utf-8');
    });


    // ============ 下面是通用的函数们 ============

    // （1）积分计算
    function calculateScores(dataRow) {
        const school_b = dataRow[1];
        const school_c = dataRow[2];
        const j_value = dataRow[7];
        const k = parseInt(dataRow[8]) || 0;
        const l = parseInt(dataRow[9]) || 0;
        const m = parseInt(dataRow[10]) || 0;
        const n_val = parseInt(dataRow[11]) || 0;
        const o_val = parseInt(dataRow[12]) || 0;
        const p = parseInt(dataRow[13]) || 0;

        let hunter_score_b, survivor_score_b, hunter_score_c, survivor_score_c;

        if (j_value === 'Yes') {
            hunter_score_b = k;
            survivor_score_b = n_val;
            hunter_score_c = o_val;
            survivor_score_c = l;
        } else if (j_value === 'No') {
            hunter_score_b = n_val;
            survivor_score_b = k;
            hunter_score_c = l;
            survivor_score_c = o_val;
        } else {
            throw new Error(`J列的值 '${j_value}' 既不是 'Yes' 也不是 'No'，请检查数据。`);
        }

        const small_score_b = hunter_score_b + survivor_score_b;
        const small_score_c = hunter_score_c + survivor_score_c;

        let winner = '';
        if (small_score_b > small_score_c) {
            winner = 'B';
        } else if (small_score_c > small_score_b) {
            winner = 'C';
        } else {
            if (j_value === 'Yes') {
                if (k === 5) {
                    if (m < p) {
                        winner = 'B';
                    } else if (m > p) {
                        winner = 'C';
                    }
                } else {
                    if (m < p) {
                        winner = 'C';
                    } else if (m > p) {
                        winner = 'B';
                    }
                }
            } else if (j_value === 'No') {
                if (l === 5) {
                    if (m < p) {
                        winner = 'C';
                    } else if (m > p) {
                        winner = 'B';
                    }
                } else {
                    if (m < p) {
                        winner = 'B';
                    } else if (m > p) {
                        winner = 'C';
                    }
                }
            } else {
                winner = 'Tie';
            }
        }

        return {
            school_b,
            school_c,
            j_value,
            hunter_score_b,
            survivor_score_b,
            hunter_score_c,
            survivor_score_c,
            small_score_b,
            small_score_c,
            winner
        };
    }

    // （2）读取 CSV 并累加数据
    function processCSV(csvData) {
        const schoolData = {};
        const rowResults = [];
        const roundData = {};

        for (let i = 1; i < csvData.length; i++) {
            const row = csvData[i];
            if (row.length < 14) {
                console.warn(`警告: CSV 文件行数据列数不足，跳过该行: ${row}`);
                continue;
            }
            if (row[4] !== "已结束") {
                continue;
            }

            try {
                const result = calculateScores(row);
                rowResults.push(result);

                const { school_b, school_c, winner, hunter_score_b, survivor_score_b, hunter_score_c, survivor_score_c } = result;

                if (!schoolData[school_b]) {
                    schoolData[school_b] = {
                        '积分': 0,
                        '小分总和': 0,
                        '求生总得分': 0,
                        '监管总得分': 0,
                        '出现次数': 0
                    };
                    roundData[school_b] = 0;
                }
                if (!schoolData[school_c]) {
                    schoolData[school_c] = {
                        '积分': 0,
                        '小分总和': 0,
                        '求生总得分': 0,
                        '监管总得分': 0,
                        '出现次数': 0
                    };
                    roundData[school_c] = 0;
                }

                roundData[school_b] += 1;
                roundData[school_c] += 1;

                if (!isNaN(result.small_score_b)) {
                    schoolData[school_b]['小分总和'] += result.small_score_b;
                }
                if (!isNaN(survivor_score_b)) {
                    schoolData[school_b]['求生总得分'] += survivor_score_b;
                }
                if (!isNaN(hunter_score_b)) {
                    schoolData[school_b]['监管总得分'] += hunter_score_b;
                }
                schoolData[school_b]['出现次数'] += 1;
                if (winner === 'B') {
                    schoolData[school_b]['积分'] += 3;
                }

                if (!isNaN(result.small_score_c)) {
                    schoolData[school_c]['小分总和'] += result.small_score_c;
                }
                if (!isNaN(survivor_score_c)) {
                    schoolData[school_c]['求生总得分'] += survivor_score_c;
                }
                if (!isNaN(hunter_score_c)) {
                    schoolData[school_c]['监管总得分'] += hunter_score_c;
                }
                schoolData[school_c]['出现次数'] += 1;
                if (winner === 'C') {
                    schoolData[school_c]['积分'] += 3;
                }

            } catch (e) {
                console.error(`处理行时出错 (学校 B: ${row[1]}, 学校 C: ${row[2]}): ${e}`);
            }
        }

        for (const schoolName in schoolData) {
            const data = schoolData[schoolName];
            const appearances = data['出现次数'];
            if (appearances > 0) {
                data['求生局均得分'] = (data['求生总得分'] / appearances).toFixed(3);
                data['监管局均得分'] = (data['监管总得分'] / appearances).toFixed(3);
            } else {
                data['求生局均得分'] = 0;
                data['监管局均得分'] = 0;
            }
        }

        return { schoolData, roundData };
    }

    // （3）对学校进行排名
    function rankSchools(schoolData) {
        const rankedSchools = [];
        for (const schoolName in schoolData) {
            const data = schoolData[schoolName];
            rankedSchools.push({
                '排名': 0,
                '学校/队伍名称': schoolName,
                '积分': data['积分'],
                '小分': data['小分总和'],
                '求生局均得分': data['求生局均得分'],
                '监管局均得分': data['监管局均得分']
            });
        }

        rankedSchools.sort((a, b) => {
            if (b['积分'] !== a['积分']) {
                return b['积分'] - a['积分'];
            }
            if (b['小分'] !== a['小分']) {
                return b['小分'] - a['小分'];
            }
            if (b['求生局均得分'] !== a['求生局均得分']) {
                return b['求生局均得分'] - a['求生局均得分'];
            }
            return b['监管局均得分'] - a['监管局均得分'];
        });

        let rank = 1;
        for (let i = 0; i < rankedSchools.length; i++) {
            if (
                i > 0 &&
                (
                    rankedSchools[i]['积分'] !== rankedSchools[i - 1]['积分'] ||
                    rankedSchools[i]['小分'] !== rankedSchools[i - 1]['小分'] ||
                    rankedSchools[i]['求生局均得分'] !== rankedSchools[i - 1]['求生局均得分'] ||
                    rankedSchools[i]['监管局均得分'] !== rankedSchools[i - 1]['监管局均得分']
                )
            ) {
                rank = i + 1;
            }
            rankedSchools[i]['排名'] = rank;
        }

        return rankedSchools;
    }

    // （4）原输出函数：直接输出积分表
    function outputRankingToExcel(rankedSchools, roundData) {
        // *** JavaScript 版本的 output_ranking_to_excel 函数 ***
        const worksheetData = [
            ["排名", "学校/队伍名称", "积分", "小分", "求生局均得分", "监管局均得分", "场次"], // Header row，添加“场次”
            ...rankedSchools.map(school => [ // 数据行
                school['排名'],
                school['学校/队伍名称'],
                school['积分'],
                school['小分'],
                parseFloat(school['求生局均得分']),
                parseFloat(school['监管局均得分']),
                roundData[school['学校/队伍名称']] // 添加场次数据
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


    // （5）新的输出函数：按组分组并输出
    function outputGroupedRankingToExcel(rankedSchools, roundData) {
        // 按组归档：你之前的分组映射
        const groupMap = {
            "Ai": [
                "北京市回民学校",
                "北京市密云区第二中学",
                "对外经济贸易大学附属中学（北京市第九十四中学）",
                "中国教育科学研究院丰台实验学校"
            ],
            "Aii": [
                "犰狳帝国",
                "北京市第三十五中学",
                "北京师范大学第二附属中学",
                "北京市海淀区教师进修学校附属实验学校"
            ],
            "Aiii": [
                "北京市信息管理学校",
                "上古神兽聚集地",
                "首都师范大学附属苹果园中学",
                "北京理工大学附属中学通州校区"
            ],
            "Aiv": [
                "北京师范大学实验中学丰台学校",
                "RXZ",
                "夙辉",
                "东北师范大学附属中学朝阳学校"
            ],
            "Bi": [
                "MXC",
                "北京大学附属中学本校",
                "北京市昌平职业学校",
                "输了怪领队"
            ],
            "Bii": [
                "北京师范大学燕化附属中学",
                "北京市通州区潞河中学",
                "北京市第十二中学",
                "Fvv"
            ],
            "Biii": [
                "QwQ",
                "北大附中朝阳未来学校",
                "北京市第二中学朝阳学校",
                "北京景山学校远洋分校"
            ],
            "Biv": [
                "中国人民大学附属中学通州校区",
                "醒",
                "中央民族大学附属中学",
                "北京市第三中学"
            ],
            "Ci": [
                "LSQY",
                "北京师范大学附属实验中学",
                "北京市第二十中学",
                "authority"
            ],
            "Cii": [
                "北京西城职业学校",
                "北京市日坛中学",
                "北京市陈经纶中学",
                "北京市第九中学"
            ],
            "Ciii": [
                "翻斗花园第一突击队",
                "北京市第二中学通州校区",
                "首都师范大学第二附属中学",
                "北京市第五中学"
            ],
            "Civ": [
                "北京理工大学附属中学",
                "北京市第一五九中学",
                "嚎叫的孤狼66"
            ]
        };
    
        // 根据 rankedSchools 做一个 {队伍名 -> 数据} 的映射
        const rankedMap = {};
        rankedSchools.forEach(item => {
            rankedMap[item["学校/队伍名称"]] = item;
        });
    
        // 准备工作表的二维数组 (AOA)
        const worksheetData = [
            ["排名", "学校/队伍名称", "积分", "小分", "求生局均得分", "监管局均得分", "场次"]
        ];
    
        // 按组输出
        for (const groupName of Object.keys(groupMap)) {
            // 先插一行：提示当前组名称
            worksheetData.push([`${groupName} 组`, null, null, null, null, null, null]);
    
            // 获取该组的队伍名数组
            const groupTeamNames = groupMap[groupName];
    
            // 如果 CSV 中没有出现某队伍，就补一个默认对象
            const groupTeamsData = groupTeamNames.map(teamName => {
                // 如果该队伍确实存在于 rankedMap，就拿它的成绩
                if (rankedMap[teamName]) {
                    return rankedMap[teamName];
                } 
                // 否则补 0 分
                return {
                    "排名": 0,               // 先占位
                    "学校/队伍名称": teamName,
                    "积分": 0,
                    "小分": 0,
                    "求生局均得分": 0,
                    "监管局均得分": 0
                };
            });
    
            // 对组内数据做排序（如果想直接用 map 顺序，删除这段即可）
            groupTeamsData.sort((a, b) => {
                if (b["积分"] !== a["积分"]) return b["积分"] - a["积分"];
                if (b["小分"] !== a["小分"]) return b["小分"] - a["小分"];
                if (parseFloat(b["求生局均得分"]) !== parseFloat(a["求生局均得分"])) {
                    return parseFloat(b["求生局均得分"]) - parseFloat(a["求生局均得分"]);
                }
                return parseFloat(b["监管局均得分"]) - parseFloat(a["监管局均得分"]);
            });
    
            // 计算组内排名
            let localRank = 1;
            for (let i = 0; i < groupTeamsData.length; i++) {
                if (
                    i > 0 &&
                    (
                        groupTeamsData[i]["积分"] !== groupTeamsData[i - 1]["积分"] ||
                        groupTeamsData[i]["小分"] !== groupTeamsData[i - 1]["小分"] ||
                        groupTeamsData[i]["求生局均得分"] !== groupTeamsData[i - 1]["求生局均得分"] ||
                        groupTeamsData[i]["监管局均得分"] !== groupTeamsData[i - 1]["监管局均得分"]
                    )
                ) {
                    localRank = i + 1;
                }
    
                const teamName = groupTeamsData[i]["学校/队伍名称"];
                const rounds = roundData[teamName] || 0; // 如果没有场次统计，就用 0
    
                worksheetData.push([
                    localRank,
                    teamName,
                    groupTeamsData[i]["积分"],
                    groupTeamsData[i]["小分"],
                    parseFloat(groupTeamsData[i]["求生局均得分"]),
                    parseFloat(groupTeamsData[i]["监管局均得分"]),
                    rounds
                ]);
            }
    
            // 每个分组后插一个空行
            worksheetData.push([null, null, null, null, null, null, null]);
        }
    
        // 底部插提示行
        const footerLine1 = "排名程序编写：TO";
        const footerLine2 = "若有数据问题请联系当场裁判";
    
        const footerRow1 = new Array(worksheetData[0].length).fill(null);
        footerRow1[0] = footerLine1;
        worksheetData.push(footerRow1);
    
        const footerRow2 = new Array(worksheetData[0].length).fill(null);
        footerRow2[0] = footerLine2;
        worksheetData.push(footerRow2);
    
        // 转换成 Sheet 并写入 Workbook
        const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "分组排名表");
    
        // 合并并设置样式：最后两行
        let mergeCell1 = {
            s: { r: worksheetData.length - 2, c: 0 },
            e: { r: worksheetData.length - 2, c: worksheetData[0].length - 1 }
        };
        if (!worksheet['!merges']) worksheet['!merges'] = [];
        worksheet['!merges'].push(mergeCell1);
    
        let cellAddress1 = XLSX.utils.encode_cell({ r: worksheetData.length - 2, c: 0 });
        if (!worksheet[cellAddress1]) worksheet[cellAddress1] = {};
        if (!worksheet[cellAddress1].s) worksheet[cellAddress1].s = {};
        worksheet[cellAddress1].s.alignment = { horizontal: "center", vertical: "center" };
        worksheet[cellAddress1].s.font = { bold: true };
    
        let mergeCell2 = {
            s: { r: worksheetData.length - 1, c: 0 },
            e: { r: worksheetData.length - 1, c: worksheetData[0].length - 1 }
        };
        worksheet['!merges'].push(mergeCell2);
    
        let cellAddress2 = XLSX.utils.encode_cell({ r: worksheetData.length - 1, c: 0 });
        if (!worksheet[cellAddress2]) worksheet[cellAddress2] = {};
        if (!worksheet[cellAddress2].s) worksheet[cellAddress2].s = {};
        worksheet[cellAddress2].s.alignment = { horizontal: "center", vertical: "center" };
        worksheet[cellAddress2].s.font = { color: { rgb: "FFFF0000" } };
    
        // 生成并触发下载
        const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'grouped_ranking_table.xlsx';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }
});