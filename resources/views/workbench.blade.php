<!DOCTYPE html>
<html lang="en" class="h-full bg-zinc-950 text-zinc-100">

<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Excel Transformation Workbench</title>
    @vite(['resources/css/app.css', 'resources/js/app.js'])
    <style>
        .monospace-grid {
            font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace;
            font-size: 0.75rem;
        }

        .custom-scroll::-webkit-scrollbar {
            width: 6px;
            height: 6px;
        }

        .custom-scroll::-webkit-scrollbar-track {
            background: #18181b;
        }

        .custom-scroll::-webkit-scrollbar-thumb {
            background: #3f3f46;
            border-radius: 3px;
        }

        .cell-changed {
            background-color: rgba(16, 185, 129, 0.15) !important;
            color: #4ade80 !important;
            border-left: 2px solid #10b981;
        }

        .cell-invalid {
            background-color: rgba(239, 68, 68, 0.15) !important;
            color: #f87171 !important;
            border-left: 2px solid #ef4444;
        }

        .cell-unchanged {
            color: #a1a1aa;
        }

        .glass-panel {
            background: rgba(24, 24, 27, 0.6);
            backdrop-filter: blur(12px);
            border: 1px solid rgba(63, 63, 70, 0.5);
        }
    </style>
</head>

<body class="h-full font-sans antialiased text-zinc-200">
    <div class="min-h-full flex flex-col">
        <!-- TOP CONTROL BAR -->
        <header
            class="border-b border-zinc-800 bg-zinc-900/80 backdrop-blur sticky top-0 z-50 px-6 py-4 flex items-center justify-between">
            <div class="flex items-center gap-3">
                <div class="h-3 w-3 rounded-full bg-emerald-500 animate-pulse"></div>
                <h1 class="text-lg font-bold tracking-tight text-white">Excel Transformation Workbench</h1>
                <span class="text-xs bg-zinc-800 text-zinc-400 px-2 py-0.5 rounded border border-zinc-700">v1.2.0
                    (Active)</span>
            </div>

            <div class="flex items-center gap-4">
                <!-- Status Notifications -->
                @if (session('success'))
                    <div id="toast-success"
                        class="text-xs text-emerald-400 bg-emerald-950/40 border border-emerald-800 px-3 py-1.5 rounded flex items-center gap-2">
                        <span class="inline-block w-1.5 h-1.5 rounded-full bg-emerald-400"></span>
                        {{ session('success') }}
                    </div>
                @endif
                @if (session('error'))
                    <div id="toast-error"
                        class="text-xs text-red-400 bg-red-950/40 border border-red-800 px-3 py-1.5 rounded flex items-center gap-2">
                        <span class="inline-block w-1.5 h-1.5 rounded-full bg-red-400"></span>
                        {{ session('error') }}
                    </div>
                @endif

                <!-- Control Buttons -->
                <form action="{{ route('workbench.upload') }}" method="POST" enctype="multipart/form-data"
                    class="flex items-center gap-2" id="upload-form">
                    @csrf
                    <label
                        class="cursor-pointer bg-zinc-800 hover:bg-zinc-700 text-white text-xs font-semibold px-3 py-1.5 rounded border border-zinc-700 transition flex items-center gap-2">
                        <svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2"
                                d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12" />
                        </svg>
                        <span>Upload File</span>
                        <input type="file" name="excel_file" class="hidden"
                            onchange="document.getElementById('upload-form').submit()">
                    </label>
                </form>

                <form action="{{ route('workbench.dry-run') }}" method="POST"
                    class="flex items-center gap-3 bg-zinc-800/50 border border-zinc-800 rounded px-3 py-1">
                    @csrf
                    <label for="strict_mode" class="text-xs text-zinc-400 select-none">Strict Mode</label>
                    <input type="checkbox" name="strict_mode" id="strict_mode"
                        class="rounded bg-zinc-950 border-zinc-700 text-emerald-500 focus:ring-0 focus:ring-offset-0"
                        onchange="this.form.submit()">
                    <button type="submit"
                        class="text-xs text-emerald-400 hover:text-emerald-300 font-semibold transition ml-2">Dry
                        Run</button>
                </form>

                @if ($fileInfo)
                    <a href="{{ route('workbench.export') }}"
                        class="bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-semibold px-3 py-1.5 rounded transition flex items-center gap-2">
                        <svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2"
                                d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
                        </svg>
                        <span>Export Cleaned</span>
                    </a>

                    <form action="{{ route('workbench.reset') }}" method="POST">
                        @csrf
                        <button type="submit"
                            class="bg-zinc-900 hover:bg-zinc-800 text-zinc-400 hover:text-white text-xs font-semibold px-3 py-1.5 rounded border border-zinc-800 transition">
                            Reset Session
                        </button>
                    </form>
                @endif
            </div>
        </header>

        <!-- MAIN LAYOUT -->
        <main class="flex-1 grid grid-cols-12 gap-6 p-6">

            <!-- LEFT PANEL: METADATA & STATUS -->
            <div class="col-span-3 flex flex-col gap-6">
                <!-- FILE INFO PANEL -->
                <div class="glass-panel rounded-lg p-4 flex flex-col gap-3">
                    <div class="flex items-center justify-between border-b border-zinc-800 pb-2">
                        <h2 class="text-xs font-semibold tracking-wider text-zinc-400 uppercase">File Metadata</h2>
                        <span class="h-2 w-2 rounded-full {{ $fileInfo ? 'bg-emerald-500' : 'bg-zinc-700' }}"></span>
                    </div>
                    @if ($fileInfo)
                        <div class="flex flex-col gap-2.5 text-xs">
                            <div>
                                <span class="text-zinc-500 block">Filename</span>
                                <span class="font-medium text-white truncate block"
                                    title="{{ $fileInfo['name'] }}">{{ $fileInfo['name'] }}</span>
                            </div>
                            <div class="grid grid-cols-2 gap-2">
                                <div>
                                    <span class="text-zinc-500 block">Total Rows</span>
                                    <span
                                        class="font-medium text-white">{{ number_format($fileInfo['total_rows']) }}</span>
                                </div>
                                <div>
                                    <span class="text-zinc-500 block">Columns Detected</span>
                                    <span class="font-medium text-white">{{ count($fileInfo['columns']) }}</span>
                                </div>
                            </div>
                            <div>
                                <span class="text-zinc-500 block">Upload Timestamp</span>
                                <span class="font-mono text-zinc-300">{{ $fileInfo['timestamp'] }}</span>
                            </div>
                            <div>
                                <span class="text-zinc-500 block">Status</span>
                                <span
                                    class="inline-flex items-center gap-1.5 text-emerald-400 font-semibold bg-emerald-950/30 px-2 py-0.5 rounded border border-emerald-900/50">
                                    {{ $fileInfo['status'] }}
                                </span>
                            </div>
                        </div>
                    @else
                        <div class="text-xs text-zinc-500 py-6 text-center">
                            No file uploaded. Use the control bar to load a dataset.
                        </div>
                    @endif
                </div>

                <!-- RULE ENGINE STATUS PANEL -->
                <div class="glass-panel rounded-lg p-4 flex flex-col gap-3">
                    <div class="flex items-center justify-between border-b border-zinc-800 pb-2">
                        <h2 class="text-xs font-semibold tracking-wider text-zinc-400 uppercase">Rule Engine
                            Configuration</h2>
                        <form action="{{ route('workbench.reload-rules') }}" method="POST">
                            @csrf
                            <button type="submit"
                                class="text-[10px] text-zinc-400 hover:text-white bg-zinc-800 hover:bg-zinc-700 px-1.5 py-0.5 rounded border border-zinc-700 transition">
                                Reload
                            </button>
                        </form>
                    </div>
                    <div class="flex flex-col gap-2.5 text-xs">
                        <div class="flex justify-between items-center">
                            <span class="text-zinc-500">Active Mappings</span>
                            <span
                                class="font-mono bg-zinc-800 px-2 py-0.5 rounded text-white">{{ count($rulesConfig['agama_to_kodtaraf'] ?? []) + count($rulesConfig['kurtawar_mapping'] ?? []) }}</span>
                        </div>
                        <div class="flex flex-col gap-1">
                            <span class="text-zinc-500">Categories</span>
                            <div class="flex flex-wrap gap-1.5 pt-1">
                                <span
                                    class="text-[10px] bg-zinc-800/80 text-zinc-300 px-2 py-0.5 rounded border border-zinc-700">Kaum
                                    &rarr; KodTaraf</span>
                                <span
                                    class="text-[10px] bg-zinc-800/80 text-zinc-300 px-2 py-0.5 rounded border border-zinc-700">KurTawar
                                    Mapping</span>
                            </div>
                        </div>
                        <div>
                            <span class="text-zinc-500 block">Rule Config Path</span>
                            <span class="font-mono text-zinc-400 text-[10px] break-all">config/excel_rules.php</span>
                        </div>
                    </div>
                </div>

                <!-- ANOMALY & VALIDATION WARNINGS PANEL -->
                <div class="glass-panel rounded-lg p-4 flex flex-col gap-3 overflow-hidden" style="max-height: 320px;">
                    <div class="flex items-center justify-between border-b border-zinc-800 pb-2">
                        <h2 class="text-xs font-semibold tracking-wider text-zinc-400 uppercase">Validation Warnings /
                            Anomalies</h2>
                        <span
                            class="bg-red-950 text-red-400 border border-red-900 text-[10px] px-2 py-0.5 rounded font-mono font-bold">{{ count($anomalies) }}</span>
                    </div>

                    <div class="overflow-y-auto custom-scroll pr-1 flex flex-col gap-2" style="max-height: 240px;">
                        @forelse($anomalies as $anomaly)
                            <div
                                class="p-2 rounded border text-xs {{ $anomaly['level'] === 'danger' ? 'bg-red-950/20 border-red-900/60 text-red-300' : 'bg-yellow-950/20 border-yellow-900/60 text-yellow-300' }}">
                                <div class="flex justify-between items-center mb-1 font-semibold">
                                    <span>Row {{ $anomaly['row'] }} &bull; {{ $anomaly['field'] }}</span>
                                    <span
                                        class="text-[9px] uppercase tracking-wider px-1 rounded {{ $anomaly['level'] === 'danger' ? 'bg-red-900/30' : 'bg-yellow-900/30' }}">
                                        {{ $anomaly['level'] }}
                                    </span>
                                </div>
                                <p class="text-[11px] leading-relaxed text-zinc-400">{{ $anomaly['description'] }}</p>
                            </div>
                        @empty
                            <div class="text-zinc-500 text-center py-12 text-xs">
                                No anomalies or validation errors detected in active dataset.
                            </div>
                        @endforelse
                    </div>
                </div>
            </div>

            <!-- RIGHT PANEL: SPLIT DATA PREVIEW WORKBENCH -->
            <div class="col-span-9 flex flex-col gap-6">

                <!-- Sheet Selector Tabs -->
                <div class="flex items-center gap-2">
                    @foreach(['PUPW', 'UTM-IDP', 'Foundation'] as $sheet)
                        <a href="?sheet={{ $sheet }}" 
                           class="px-4 py-1.5 text-xs font-semibold rounded border transition {{ ($activeSheet ?? 'PUPW') === $sheet ? 'bg-emerald-600 border-emerald-500 text-white font-bold' : 'bg-zinc-900/50 border-zinc-800 text-zinc-400 hover:text-white' }}">
                            {{ $sheet }}
                        </a>
                    @endforeach
                </div>

                <!-- TABLE VIEWPORT CONTROLS -->
                <div
                    class="flex items-center justify-between bg-zinc-900/50 border border-zinc-800 px-4 py-2 rounded-lg">
                    <div class="flex items-center gap-4">
                        <div class="flex items-center gap-2">
                            <span class="h-2 w-2 rounded-full bg-emerald-500"></span>
                            <span class="text-xs text-zinc-400">Green = Transformed</span>
                        </div>
                        <div class="flex items-center gap-2">
                            <span class="h-2 w-2 rounded-full bg-red-500"></span>
                            <span class="text-xs text-zinc-400">Red = Mapping Error</span>
                        </div>
                        <div class="flex items-center gap-2">
                            <span class="h-2 w-2 rounded-full bg-zinc-500"></span>
                            <span class="text-xs text-zinc-400">Grey = Unchanged</span>
                        </div>
                    </div>

                    <div class="flex items-center gap-3">
                        <label class="inline-flex items-center gap-2 text-xs cursor-pointer select-none">
                            <input type="checkbox" id="toggle-changed"
                                class="rounded bg-zinc-950 border-zinc-800 text-emerald-500 focus:ring-0 focus:ring-offset-0">
                            <span class="text-zinc-400">Show only changed rows</span>
                        </label>
                        <input type="text" id="row-search" placeholder="Search rows..."
                            class="bg-zinc-950 border border-zinc-800 rounded px-2 py-1 text-xs text-zinc-200 placeholder-zinc-600 focus:outline-none focus:border-zinc-700 w-48">
                    </div>
                </div>

                <!-- SIDE-BY-SIDE GRID PREVIEW -->
                <div class="flex-1 grid grid-cols-2 gap-4 min-h-0">

                    <!-- LEFT COLUMN: RAW EXCEL PREVIEW -->
                    <div class="glass-panel rounded-lg flex flex-col min-h-0">
                        <div
                            class="px-4 py-2.5 border-b border-zinc-800 bg-zinc-900/40 flex items-center justify-between">
                            <span class="text-xs font-semibold text-zinc-300">Raw Excel Input (First 20 Rows)</span>
                            <span class="text-[10px] text-zinc-500">Original Values</span>
                        </div>
                        <!-- Top Fake Scrollbar for Raw Table -->
                        <div class="overflow-x-auto custom-scroll w-full border-b border-zinc-900 bg-zinc-950/40"
                            id="raw-fake-scroll" style="height: 6px; display: none;">
                            <div id="raw-fake-content" style="height: 1px;"></div>
                        </div>
                        <div class="flex-1 overflow-auto custom-scroll monospace-grid" id="raw-scroll-container"
                            style="max-height: 600px;">
                            <table class="w-full text-left border-collapse" id="raw-table">
                                <thead>
                                    <tr class="bg-zinc-900 border-b border-zinc-800 text-zinc-400 sticky top-0 z-10">
                                        <th class="p-2 border-r border-zinc-800 w-12 text-center">Row</th>
                                        @if ($fileInfo)
                                            @foreach ($fileInfo['columns'] as $col)
                                                <th class="p-2 border-r border-zinc-800 min-w-[120px]">
                                                    {{ $col }}</th>
                                            @endforeach
                                        @else
                                            <th class="p-2">Data Columns</th>
                                        @endif
                                    </tr>
                                </thead>
                                <tbody>
                                    @forelse($rawRows as $rowNum => $row)
                                        <tr class="border-b border-zinc-900 hover:bg-zinc-900/30 transition-colors row-item"
                                            data-row-num="{{ $rowNum }}">
                                            <td
                                                class="p-2 border-r border-zinc-800 bg-zinc-900/40 text-center font-semibold text-zinc-500">
                                                {{ $rowNum }}</td>
                                            @foreach ($fileInfo['columns'] as $col)
                                                @php
                                                    $val = $row[$col] ?? '';
                                                    $isKaum = $col === 'ASA_KAUM';
                                                    $isKurTawar = $col === 'ASA_KURTAWAR';
                                                    $cellClass = 'cell-unchanged';
                                                    if ($isKaum && $val !== '') {
                                                        $cellClass = 'cell-changed';
                                                    }
                                                    if ($isKurTawar && $val !== '') {
                                                        $mapping = config('excel_rules.kurtawar_mapping');
                                                        $cellClass = isset($mapping[$val])
                                                            ? 'cell-changed'
                                                            : 'cell-invalid';
                                                    }
                                                @endphp
                                                <td class="p-2 border-r border-zinc-800 {{ $cellClass }}">
                                                    {{ $val !== '' ? $val : '-' }}</td>
                                            @endforeach
                                        </tr>
                                    @empty
                                        <tr>
                                            <td colspan="10" class="p-8 text-center text-zinc-600">No Excel file
                                                loaded.</td>
                                        </tr>
                                    @endforelse
                                </tbody>
                            </table>
                        </div>
                    </div>

                    <!-- RIGHT COLUMN: TRANSFORMED PREVIEW -->
                    <div class="glass-panel rounded-lg flex flex-col min-h-0">
                        <div
                            class="px-4 py-2.5 border-b border-zinc-800 bg-zinc-900/40 flex items-center justify-between">
                            <span class="text-xs font-semibold text-zinc-300">Transformed Preview Output</span>
                            <span class="text-[10px] text-zinc-500">Processed Values</span>
                        </div>
                        <!-- Top Fake Scrollbar for Transformed Table -->
                        <div class="overflow-x-auto custom-scroll w-full border-b border-zinc-900 bg-zinc-950/40"
                            id="transformed-fake-scroll" style="height: 6px; display: none;">
                            <div id="transformed-fake-content" style="height: 1px;"></div>
                        </div>
                        <div class="flex-1 overflow-auto custom-scroll monospace-grid"
                            id="transformed-scroll-container" style="max-height: 600px;">
                            <table class="w-full text-left border-collapse" id="transformed-table">
                                <thead>
                                    <tr class="bg-zinc-900 border-b border-zinc-800 text-zinc-400 sticky top-0 z-10">
                                        <th class="p-2 border-r border-zinc-800 w-12 text-center">Row</th>
                                        @if ($fileInfo)
                                            <!-- List all columns, plus generated columns if not already in headers -->
                                            @php
                                                $outCols = $fileInfo['columns'];
                                                if (!in_array('ASA_KODTARAF', $outCols)) {
                                                    $outCols[] = 'ASA_KODTARAF';
                                                }
                                            @endphp
                                            @foreach ($outCols as $col)
                                                <th class="p-2 border-r border-zinc-800 min-w-[120px]">
                                                    {{ $col }}</th>
                                            @endforeach
                                        @else
                                            <th class="p-2">Processed Columns</th>
                                        @endif
                                    </tr>
                                </thead>
                                <tbody>
                                    @forelse($transformedRows as $rowNum => $row)
                                        <tr class="border-b border-zinc-900 hover:bg-zinc-900/30 transition-colors row-item"
                                            data-row-num="{{ $rowNum }}">
                                            <td
                                                class="p-2 border-r border-zinc-800 bg-zinc-900/40 text-center font-semibold text-zinc-500">
                                                {{ $rowNum }}</td>
                                            @php
                                                $outCols = $fileInfo['columns'];
                                                if (!in_array('ASA_KODTARAF', $outCols)) {
                                                    $outCols[] = 'ASA_KODTARAF';
                                                }
                                            @endphp
                                            @foreach ($outCols as $col)
                                                @php
                                                    $val = $row[$col] ?? '';
                                                    $rawVal = $rawRows[$rowNum][$col] ?? null;

                                                    // Determine visual status dynamically based on diff
                                                    $cellClass = 'cell-unchanged';
                                                    if ($rawVal !== null && $val !== $rawVal) {
                                                        $cellClass = 'cell-changed';
                                                    }
                                                    // Highlight unmapped KurTawar codes as invalid
                                                    if ($col === 'ASA_KURTAWAR') {
                                                        $rawKur = $rawRows[$rowNum]['ASA_KURTAWAR'] ?? '';
                                                        if ($rawKur !== '' && !isset(config('excel_rules.kurtawar_mapping')[$rawKur])) {
                                                            $cellClass = 'cell-invalid';
                                                        }
                                                    }
                                                @endphp
                                                <td class="p-2 border-r border-zinc-800 {{ $cellClass }}">
                                                    {{ $val !== '' ? $val : '-' }}</td>
                                            @endforeach
                                        </tr>
                                    @empty
                                        <tr>
                                            <td colspan="10" class="p-8 text-center text-zinc-600">No Excel file
                                                loaded.</td>
                                        </tr>
                                    @endforelse
                                </tbody>
                            </table>
                        </div>
                    </div>

                </div>

                <!-- BOTTOM ACTION BAR -->
                <div class="flex items-center justify-between border-t border-zinc-800 pt-4">
                    <div class="text-xs text-zinc-500">
                        @if ($fileInfo)
                            Displaying 1 - {{ min($fileInfo['total_rows'], 20) }} of
                            {{ number_format($fileInfo['total_rows']) }} rows
                        @else
                            Waiting for dataset upload...
                        @endif
                    </div>

                    @if ($fileInfo)
                        <div class="flex items-center gap-3">
                            <form action="{{ route('workbench.reload-rules') }}" method="POST">
                                @csrf
                                <button type="submit"
                                    class="bg-zinc-900 hover:bg-zinc-800 text-zinc-300 text-xs font-semibold px-4 py-2 rounded border border-zinc-800 transition">
                                    Re-run with Updated Rules
                                </button>
                            </form>
                            <a href="{{ route('workbench.export') }}"
                                class="bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-semibold px-4 py-2 rounded transition flex items-center gap-2">
                                <span>Export excel output</span>
                            </a>
                        </div>
                    @endif
                </div>

            </div>

        </main>
    </div>

    <!-- Sync scroll logic -->
    <script>
        const rawScroll = document.getElementById('raw-scroll-container');
        const transScroll = document.getElementById('transformed-scroll-container');
        const rawFakeScroll = document.getElementById('raw-fake-scroll');
        const rawFakeContent = document.getElementById('raw-fake-content');
        const rawTable = document.getElementById('raw-table');
        const transFakeScroll = document.getElementById('transformed-fake-scroll');
        const transFakeContent = document.getElementById('transformed-fake-content');
        const transTable = document.getElementById('transformed-table');

        function initFakeScrollbars() {
            if (rawTable && rawFakeScroll && rawFakeContent && rawScroll) {
                const tableWidth = rawTable.offsetWidth;
                const containerWidth = rawScroll.offsetWidth;
                if (tableWidth > containerWidth) {
                    rawFakeScroll.style.display = 'block';
                    rawFakeContent.style.width = tableWidth + 'px';
                } else {
                    rawFakeScroll.style.display = 'none';
                }
            }

            if (transTable && transFakeScroll && transFakeContent && transScroll) {
                const tableWidth = transTable.offsetWidth;
                const containerWidth = transScroll.offsetWidth;
                if (tableWidth > containerWidth) {
                    transFakeScroll.style.display = 'block';
                    transFakeContent.style.width = tableWidth + 'px';
                } else {
                    transFakeScroll.style.display = 'none';
                }
            }
        }

        // Run scroll synchronizations
        if (rawScroll && transScroll) {
            let isSyncingRawScroll = false;
            let isSyncingTransScroll = false;

            rawScroll.onscroll = function() {
                if (!isSyncingRawScroll) {
                    isSyncingTransScroll = true;
                    transScroll.scrollTop = this.scrollTop;
                    transScroll.scrollLeft = this.scrollLeft;
                    if (transFakeScroll) transFakeScroll.scrollLeft = this.scrollLeft;
                    if (rawFakeScroll) rawFakeScroll.scrollLeft = this.scrollLeft;
                }
                isSyncingRawScroll = false;
            };

            transScroll.onscroll = function() {
                if (!isSyncingTransScroll) {
                    isSyncingRawScroll = true;
                    rawScroll.scrollTop = this.scrollTop;
                    rawScroll.scrollLeft = this.scrollLeft;
                    if (rawFakeScroll) rawFakeScroll.scrollLeft = this.scrollLeft;
                    if (transFakeScroll) transFakeScroll.scrollLeft = this.scrollLeft;
                }
                isSyncingTransScroll = false;
            };
        }

        // Sync top fake scrolls with tables
        if (rawFakeScroll && rawScroll) {
            let isSyncingFake = false;
            rawFakeScroll.onscroll = function() {
                if (!isSyncingFake) {
                    isSyncingFake = true;
                    rawScroll.scrollLeft = this.scrollLeft;
                }
                isSyncingFake = false;
            };
        }

        if (transFakeScroll && transScroll) {
            let isSyncingFake = false;
            transFakeScroll.onscroll = function() {
                if (!isSyncingFake) {
                    isSyncingFake = true;
                    transScroll.scrollLeft = this.scrollLeft;
                }
                isSyncingFake = false;
            };
        }

        window.addEventListener('load', initFakeScrollbars);
        window.addEventListener('resize', initFakeScrollbars);
        setTimeout(initFakeScrollbars, 500);

        // Live Row Search
        const searchInput = document.getElementById('row-search');
        if (searchInput) {
            searchInput.addEventListener('input', function() {
                const query = this.value.toLowerCase();
                document.querySelectorAll('.row-item').forEach(row => {
                    const text = row.innerText.toLowerCase();
                    const rowNum = row.getAttribute('data-row-num');
                    const pairRows = document.querySelectorAll(`[data-row-num="${rowNum}"]`);

                    if (text.includes(query)) {
                        pairRows.forEach(r => r.style.display = '');
                    } else {
                        pairRows.forEach(r => r.style.display = 'none');
                    }
                });
                initFakeScrollbars();
            });
        }

        // Show only changed rows filter
        const toggleChanged = document.getElementById('toggle-changed');
        if (toggleChanged) {
            toggleChanged.addEventListener('change', function() {
                const checked = this.checked;
                document.querySelectorAll('.row-item').forEach(row => {
                    const rowNum = row.getAttribute('data-row-num');
                    const pairRows = document.querySelectorAll(`[data-row-num="${rowNum}"]`);

                    let hasChange = false;
                    pairRows.forEach(r => {
                        if (r.querySelector('.cell-changed') || r.querySelector('.cell-invalid')) {
                            hasChange = true;
                        }
                    });

                    if (checked) {
                        if (hasChange) {
                            pairRows.forEach(r => r.style.display = '');
                        } else {
                            pairRows.forEach(r => r.style.display = 'none');
                        }
                    } else {
                        // respect search query if any
                        const query = searchInput ? searchInput.value.toLowerCase() : '';
                        pairRows.forEach(r => {
                            const text = r.innerText.toLowerCase();
                            if (text.includes(query)) {
                                r.style.display = '';
                            }
                        });
                    }
                });
                initFakeScrollbars();
            });
        }
    </script>
</body>

</html>
