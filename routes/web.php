<?php

use App\Http\Controllers\ExcelController;
use Illuminate\Support\Facades\Route;

Route::get('/', [ExcelController::class, 'index'])->name('workbench.index');
Route::post('/upload', [ExcelController::class, 'upload'])->name('workbench.upload');
Route::post('/dry-run', [ExcelController::class, 'dryRun'])->name('workbench.dry-run');
Route::post('/reload-rules', [ExcelController::class, 'reloadRules'])->name('workbench.reload-rules');
Route::get('/export', [ExcelController::class, 'export'])->name('workbench.export');
Route::post('/reset', [ExcelController::class, 'reset'])->name('workbench.reset');




// use Illuminate\Http\Request;
// use Illuminate\Support\Facades\Route;
// use App\Http\Controllers\ExcelController;

// Route::prefix('workbench')->name('workbench.')->group(function () {
//     Route::get('/', [ExcelController::class, 'index'])->name('index');
//     Route::post('/upload', [ExcelController::class, 'upload'])->name('upload');
//     Route::post('/dry-run', [ExcelController::class, 'dryRun'])->name('dry-run');
//     Route::post('/reload-rules', [ExcelController::class, 'reloadRules'])->name('reload-rules');
//     Route::get('/export', [ExcelController::class, 'export'])->name('export');
//     Route::post('/reset', [ExcelController::class, 'reset'])->name('reset');
// });