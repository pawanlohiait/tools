# tools
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Excel Value Finder Pro | Created by Pawan Lohiya</title>
    <!-- Tailwind CSS -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- SheetJS Library -->
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
    <!-- FontAwesome -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
</head>
<body class="bg-slate-100 text-slate-800 min-h-screen flex flex-col font-sans">

    <!-- Top Navigation Bar -->
    <header class="bg-white border-b border-slate-200 sticky top-0 z-30 shadow-sm">
        <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-3.5 flex flex-col sm:flex-row justify-between items-center gap-3">
            <div class="flex items-center space-x-3 w-full sm:w-auto justify-between sm:justify-start">
                <div class="flex items-center space-x-3">
                    <div class="w-10 h-10 rounded-xl bg-emerald-600 flex items-center justify-center text-white shadow-md shadow-emerald-600/20 flex-shrink-0">
                        <i class="fa-solid fa-file-excel text-lg"></i>
                    </div>
                    <div>
                        <h1 class="text-base sm:text-lg font-bold text-slate-900 leading-tight">Excel &amp; WPS Value Finder Pro</h1>
                        <p class="text-xs text-slate-500">Created by <span class="text-slate-900 font-bold">Pawan Lohiya</span></p>
                    </div>
                </div>
            </div>
            <button onclick="loadSampleData()" class="w-full sm:w-auto px-4 py-2 bg-emerald-600 hover:bg-emerald-700 text-xs font-bold text-white rounded-xl transition cursor-pointer shadow-sm flex items-center justify-center gap-2">
                <i class="fa-solid fa-wand-magic-sparkles"></i> Load Sample Data
            </button>
        </div>
    </header>

    <!-- Main Responsive Container -->
    <main class="flex-grow max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-6 sm:py-8 space-y-6">

        <!-- Mode Selector Tabs -->
        <div class="flex gap-2 border-b border-slate-300 pb-2">
            <button onclick="switchTab('upload')" id="tabUploadBtn" class="px-4 py-2 bg-emerald-600 text-white text-xs font-bold rounded-xl transition cursor-pointer shadow-xs">
                📂 1. File Upload Mode
            </button>
            <button onclick="switchTab('paste')" id="tabPasteBtn" class="px-4 py-2 bg-white text-slate-700 border border-slate-300 text-xs font-bold rounded-xl transition cursor-pointer hover:bg-slate-50">
                📋 2. Direct Table Data Paste (Backup)
            </button>
        </div>

        <!-- Desktop Responsive Grid for Upload/Paste & Search Options -->
        <div class="grid grid-cols-1 lg:grid-cols-12 gap-6">

            <!-- Left Panel: Upload or Paste (5 cols on desktop) -->
            <div class="lg:col-span-5 space-y-6 flex flex-col">
                
                <!-- TAB 1: File Upload Box -->
                <div id="sectionUpload" class="bg-white p-6 rounded-2xl border border-slate-200 space-y-4 shadow-sm flex flex-col flex-grow">
                    <h2 class="text-sm font-bold text-slate-900 flex items-center gap-2">
                        <span class="w-6 h-6 rounded-full bg-emerald-100 text-emerald-700 flex items-center justify-center text-xs font-bold">1</span>
                        Upload Your Spreadsheets (.xlsx, .et, .csv)
                    </h2>
                    
                    <div class="border-2 border-dashed border-slate-300 hover:border-emerald-500 rounded-xl p-6 text-center cursor-pointer bg-slate-50/50 transition-all flex flex-col items-center justify-center flex-grow min-h-[160px]" onclick="document.getElementById('fileInput').click()">
                        <input type="file" id="fileInput" multiple accept=".xlsx, .xls, .csv, .et, .ett, .ods" class="hidden" onchange="handleFileSelect(event)">
                        <div class="w-12 h-12 rounded-full bg-emerald-50 text-emerald-600 flex items-center justify-center mb-3 shadow-xs">
                            <i class="fa-solid fa-cloud-arrow-up text-xl"></i>
                        </div>
                        <p class="text-sm font-semibold text-slate-800">Click to browse or drag &amp; drop files</p>
                        <p class="text-xs text-slate-500 mt-1">Supports WPS (.et), Excel (.xlsx), &amp; CSV</p>
                    </div>
                    <div id="fileStatus" class="text-xs text-emerald-700 font-mono font-medium bg-emerald-50/60 p-3 rounded-xl border border-emerald-100">No files uploaded yet.</div>
                </div>

                <!-- TAB 2: Direct Paste Mode -->
                <div id="sectionPaste" class="bg-white p-6 rounded-2xl border border-slate-200 space-y-4 shadow-sm hidden flex flex-col flex-grow">
                    <h2 class="text-sm font-bold text-slate-900 flex items-center gap-2">
                        <span class="w-6 h-6 rounded-full bg-emerald-100 text-emerald-700 flex items-center justify-center text-xs">💡</span>
                        Paste Excel Table Data Directly
                    </h2>
                    <p class="text-xs text-slate-500">If your browser blocks local files, copy rows from Excel/WPS (<kbd class="bg-slate-100 px-1.5 py-0.5 rounded text-emerald-700 font-bold">Ctrl+C</kbd>) and paste below:</p>
                    
                    <div class="space-y-1">
                        <label class="text-xs font-semibold text-slate-700">File Name Label:</label>
                        <input type="text" id="pasteFileName" placeholder="e.g., Sales_Data.xlsx" class="w-full bg-slate-50 border border-slate-300 rounded-xl p-2.5 text-xs text-slate-900 focus:outline-none focus:border-emerald-500 font-mono">
                    </div>

                    <div class="space-y-1 flex-grow flex flex-col">
                        <label class="text-xs font-semibold text-slate-700">Paste Table Rows (Tab-separated):</label>
                        <textarea id="pasteDataBox" rows="4" placeholder="Paste copied rows from Excel here..." class="w-full bg-slate-50 border border-slate-300 rounded-xl p-3 text-xs text-slate-900 focus:outline-none focus:border-emerald-500 font-mono resize-y flex-grow"></textarea>
                    </div>

                    <button onclick="savePastedData()" class="w-full py-2.5 bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-xs rounded-xl transition cursor-pointer shadow-xs">
                        ➕ Index This Pasted Data
                    </button>
                </div>

            </div>

            <!-- Right Panel: Search Box & Advanced Options (7 cols on desktop) -->
            <div class="lg:col-span-7 flex flex-col">
                <div class="bg-white p-6 rounded-2xl border border-slate-200 space-y-5 shadow-sm flex flex-col h-full justify-between">
                    <div class="space-y-4">
                        <h2 class="text-sm font-bold text-slate-900 flex items-center gap-2">
                            <span class="w-6 h-6 rounded-full bg-emerald-100 text-emerald-700 flex items-center justify-center text-xs font-bold">2</span>
                            Enter Values to Search (Includes Row 1 / Headers)
                        </h2>
                        
                        <textarea id="searchBox" rows="5" placeholder="Enter values to search (one per line):&#10;ORD-1001&#10;Pawan Lohiya&#10;a" class="w-full bg-slate-50 border border-slate-300 rounded-xl p-3.5 text-xs text-slate-900 focus:outline-none focus:border-emerald-500 font-mono resize-y shadow-inner"></textarea>
                        
                        <!-- Advanced Search Options Bar -->
                        <div class="grid grid-cols-1 sm:grid-cols-2 gap-3 pt-3 border-t border-slate-100 text-xs">
                            <label class="flex items-center gap-2.5 cursor-pointer font-semibold text-slate-700 select-none bg-slate-50 p-3 rounded-xl border border-slate-200 hover:bg-slate-100 transition">
                                <input type="checkbox" id="exactMatchCheck" class="w-4 h-4 text-emerald-600 rounded border-slate-300 focus:ring-emerald-500 cursor-pointer flex-shrink-0">
                                <span>Exact Match Only (e.g. "a" won't match "apple")</span>
                            </label>
                            
                            <label class="flex items-center gap-2.5 cursor-pointer font-semibold text-slate-700 select-none bg-slate-50 p-3 rounded-xl border border-slate-200 hover:bg-slate-100 transition">
                                <input type="checkbox" id="caseSensitiveCheck" class="w-4 h-4 text-emerald-600 rounded border-slate-300 focus:ring-emerald-500 cursor-pointer flex-shrink-0">
                                <span>Case-Sensitive Match</span>
                            </label>
                        </div>
                    </div>

                    <button onclick="executeSearch()" class="w-full py-3.5 bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-sm rounded-xl transition cursor-pointer shadow-md flex items-center justify-center gap-2 mt-4">
                        <i class="fa-solid fa-magnifying-glass"></i> Find Which Excel Has These Values
                    </button>
                </div>
            </div>

        </div>

        <!-- Results Section (Full Width Responsive Table) -->
        <div id="resultsCard" class="bg-white p-6 rounded-2xl border border-slate-200 hidden space-y-4 shadow-sm">
            <h2 class="text-sm font-bold text-slate-900 flex items-center gap-2">
                <i class="fa-solid fa-clipboard-list text-emerald-600"></i> Search Results
            </h2>
            <div class="overflow-x-auto rounded-xl border border-slate-200">
                <table class="w-full text-left text-xs font-mono min-w-[600px]">
                    <thead>
                        <tr class="bg-slate-100 text-slate-700 border-b border-slate-200">
                            <th class="p-3.5">Search Value</th>
                            <th class="p-3.5">Status</th>
                            <th class="p-3.5">File Name</th>
                            <th class="p-3.5">Sheet Name</th>
                            <th class="p-3.5">Cell Location</th>
                        </tr>
                    </thead>
                    <tbody id="resultsTableBody" class="divide-y divide-slate-200 text-slate-800"></tbody>
                </table>
            </div>
        </div>

    </main>

    <!-- Footer -->
    <footer class="bg-white border-t border-slate-200 mt-auto py-4 text-center text-xs text-slate-500">
        Excel &amp; WPS Value Finder Pro &bull; Created with ❤️ by <span class="font-bold text-slate-800">Pawan Lohiya</span>
    </footer>

    <!-- Script Logic -->
    <script>
        let globalWorkbooks = {};

        function switchTab(tab) {
            if (tab === 'upload') {
                document.getElementById('sectionUpload').classList.remove('hidden');
                document.getElementById('sectionPaste').classList.add('hidden');
                document.getElementById('tabUploadBtn').className = "px-4 py-2 bg-emerald-600 text-white text-xs font-bold rounded-xl transition cursor-pointer shadow-xs";
                document.getElementById('tabPasteBtn').className = "px-4 py-2 bg-white text-slate-700 border border-slate-300 text-xs font-bold rounded-xl transition cursor-pointer hover:bg-slate-50";
            } else {
                document.getElementById('sectionUpload').classList.add('hidden');
                document.getElementById('sectionPaste').classList.remove('hidden');
                document.getElementById('tabPasteBtn').className = "px-4 py-2 bg-emerald-600 text-white text-xs font-bold rounded-xl transition cursor-pointer shadow-xs";
                document.getElementById('tabUploadBtn').className = "px-4 py-2 bg-white text-slate-700 border border-slate-300 text-xs font-bold rounded-xl transition cursor-pointer hover:bg-slate-50";
            }
        }

        async function handleFileSelect(event) {
            const files = event.target.files;
            if (!files || files.length === 0) return;

            let loadedNames = [];
            for (let file of files) {
                try {
                    const arrayBuffer = await file.arrayBuffer();
                    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
                    
                    let sheetsData = {};
                    workbook.SheetNames.forEach(sheetName => {
                        const worksheet = workbook.Sheets[sheetName];
                        const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
                        sheetsData[sheetName] = rows;
                    });

                    globalWorkbooks[file.name] = sheetsData;
                    loadedNames.push(file.name);
                } catch (err) {
                    console.error(err);
                    alert(`Error reading "${file.name}". Try using the "Direct Table Data Paste" tab above!`);
                }
            }

            document.getElementById('fileStatus').textContent = `Successfully loaded: ${loadedNames.join(', ')}`;
        }

        function savePastedData() {
            const fileName = document.getElementById('pasteFileName').value.trim() || "Pasted_Spreadsheet.xlsx";
            const rawText = document.getElementById('pasteDataBox').value.trim();

            if (!rawText) {
                alert('Please paste some rows from your Excel sheet first.');
                return;
            }

            const lines = rawText.split('\n');
            const rows = lines.map(line => line.split('\t').map(cell => cell.trim()));

            globalWorkbooks[fileName] = { "Sheet1": rows };
            document.getElementById('fileStatus').textContent = `Successfully indexed pasted file: ${fileName}`;
            alert(`Successfully indexed "${fileName}" from pasted text!`);
            switchTab('upload');
        }

        function executeSearch() {
            if (Object.keys(globalWorkbooks).length === 0) {
                alert('Please upload your files or paste data first!');
                return;
            }

            const searchInput = document.getElementById('searchBox').value.trim();
            if (!searchInput) {
                alert('Please enter at least one value to search.');
                return;
            }

            const exactMatchOnly = document.getElementById('exactMatchCheck').checked;
            const caseSensitive = document.getElementById('caseSensitiveCheck').checked;
            const queries = searchInput.split('\n').map(v => v.trim()).filter(v => v);
            let tableHtml = '';

            queries.forEach(query => {
                let foundList = [];

                for (let fileName in globalWorkbooks) {
                    const sheets = globalWorkbooks[fileName];
                    for (let sheetName in sheets) {
                        const rows = sheets[sheetName];
                        
                        for (let r = 0; r < rows.length; r++) {
                            for (let c = 0; c < rows[r].length; c++) {
                                const cellVal = String(rows[r][c] || '').trim();
                                
                                let isMatch = false;
                                if (exactMatchOnly) {
                                    if (caseSensitive) {
                                        isMatch = (cellVal === query);
                                    } else {
                                        isMatch = (cellVal.toLowerCase() === query.toLowerCase());
                                    }
                                } else {
                                    if (caseSensitive) {
                                        isMatch = cellVal.includes(query);
                                    } else {
                                        isMatch = cellVal.toLowerCase().includes(query.toLowerCase());
                                    }
                                }

                                if (isMatch) {
                                    const colLetter = XLSX.utils.encode_col(c);
                                    foundList.push({
                                        fileName: fileName,
                                        sheetName: sheetName,
                                        cell: `${colLetter}${r + 1}`
                                    });
                                }
                            }
                        }
                    }
                }

                if (foundList.length > 0) {
                    foundList.forEach(match => {
                        tableHtml += `
                            <tr class="hover:bg-slate-50">
                                <td class="p-3.5 font-bold text-slate-900">${escapeHtml(query)}</td>
                                <td class="p-3.5"><span class="px-2.5 py-0.5 rounded-full bg-emerald-100 text-emerald-800 font-bold text-[11px]">Found</span></td>
                                <td class="p-3.5 text-emerald-700 font-semibold">${escapeHtml(match.fileName)}</td>
                                <td class="p-3.5 text-slate-600">${escapeHtml(match.sheetName)}</td>
                                <td class="p-3.5 text-amber-700 font-bold">${match.cell}</td>
                            </tr>
                        `;
                    });
                } else {
                    tableHtml += `
                        <tr class="hover:bg-slate-50">
                            <td class="p-3.5 font-bold text-slate-500">${escapeHtml(query)}</td>
                            <td class="p-3.5"><span class="px-2.5 py-0.5 rounded-full bg-rose-100 text-rose-800 font-bold text-[11px]">Not Found</span></td>
                            <td class="p-3.5 text-slate-400">-</td>
                            <td class="p-3.5 text-slate-400">-</td>
                            <td class="p-3.5 text-slate-400">-</td>
                        </tr>
                    `;
                }
            });

            document.getElementById('resultsTableBody').innerHTML = tableHtml;
            document.getElementById('resultsCard').classList.remove('hidden');
            document.getElementById('resultsCard').scrollIntoView({ behavior: 'smooth' });
        }

        function escapeHtml(str) {
            return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
        }

        function loadSampleData() {
            const wb = XLSX.utils.book_new();
            const ws = XLSX.utils.aoa_to_sheet([
                ["a", "CUSTOMER", "CITY"],
                ["ORD-1001", "Pawan Lohiya", "Haridwar"],
                ["a", "Aman Kumar", "Dehradun"]
            ]);
            XLSX.utils.book_append_sheet(wb, ws, "Sales");

            const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
            const fakeFile = new File([wbout], "Sample_Worksheet.xlsx", { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });

            let dt = new DataTransfer();
            dt.items.add(fakeFile);
            document.getElementById('fileInput').files = dt.files;
            
            handleFileSelect({ target: document.getElementById('fileInput') });
            document.getElementById('searchBox').value = "a";
            document.getElementById('exactMatchCheck').checked = true;
        }
    </script>
</body>
</html>
