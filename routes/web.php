<?php

use Illuminate\Support\Facades\Route;
use App\Http\Controllers\ResetPassword;
use App\Http\Controllers\ChangePassword;
use App\Http\Controllers\PageController;
use App\Http\Controllers\UserController;
use App\Http\Controllers\LoginController;
use App\Http\Controllers\RegisterController;
use App\Http\Controllers\admin\adminController;
use App\Http\Controllers\tiket\tiketController;
use App\Http\Controllers\UserProfileController;

/*
|--------------------------------------------------------------------------
| Web Routes
|--------------------------------------------------------------------------
|
| Here is where you can register web routes for your application. These
| routes are loaded by the RouteServiceProvider within a group which
| contains the "web" middleware group. Now create something great!
|
*/

// Route::get('/', function () {
//     return view('welcome');
// });

Route::get('/', function () {
    return redirect('/dashboard');
})->middleware('auth');
Route::get('/register', [RegisterController::class, 'create'])->middleware('guest')->name('register');
Route::post('/register', [RegisterController::class, 'store'])->middleware('guest')->name('register.perform');
Route::get('/login', [LoginController::class, 'show'])->middleware('guest')->name('login');
Route::post('/login', [LoginController::class, 'login'])->middleware('guest')->name('login.perform');
Route::get('/reset-password', [ResetPassword::class, 'show'])->middleware('guest')->name('reset-password');
Route::post('/reset-password', [ResetPassword::class, 'send'])->middleware('guest')->name('reset.perform');
Route::get('/change-password', [ChangePassword::class, 'show'])->middleware('guest')->name('change-password');
Route::post('/change-password', [ChangePassword::class, 'update'])->middleware('guest')->name('change.perform');
// tiket user
Route::get('/dashboard', [tiketController::class, 'index'])->name('home')->middleware('auth');
Route::get('dashboard/Edit/{id}/', [tiketController::class, 'edit'])->name('edit')->middleware('auth');
Route::post('dashboard/Store', [tiketController::class, 'store'])->name('store')->middleware('auth');
Route::get('dashboard/Delete/{id}', [tiketController::class, 'destroy'])->name('destroy')->middleware('auth');
Route::get('dashboard/Done', [tiketController::class, 'done'])->name('done')->middleware('auth');
Route::get('dashboard/Detail/{id}', [tiketController::class, 'showDetail'])->name('show')->middleware('auth');
Route::get('/dashboard/send_email/{id}', [tiketController::class, 'send_email'])->name('send_email');
// tiket admin
Route::get('/admin', [adminController::class, 'index'])->name('admin')->middleware('auth');
Route::get('admin/Edit/{id}/', [adminController::class, 'edit'])->name('admin-edit')->middleware('auth');
Route::post('admin/Store', [adminController::class, 'store'])->name('admin-add')->middleware('auth');
Route::get('admin/Delete/{id}', [adminController::class, 'destroy'])->name('admin-delete')->middleware('auth');
Route::get('admin/Done', [adminController::class, 'done'])->name('admin-done')->middleware('auth');
Route::get('admin/Detail/{id}', [adminController::class, 'showDetail'])->name('admin-detail')->middleware('auth');
// Route User
Route::get('user', [UserController::class, 'index'])->name('user')->middleware('auth');
Route::get('user/Edit/{id}/', [UserController::class, 'edit'])->name('user.edit')->middleware('auth');
Route::post('user/Store/', [UserController::class, 'store'])->name('user.store')->middleware('auth');
Route::get('user/Detail/{id}', [UserController::class, 'showDetail'])->middleware('auth');
Route::get('user/Delete/{id}', [UserController::class, 'destroy'])->middleware('auth');
// PDF
Route::get('generate-pdf', [adminController::class, 'generatePDF'])->name('generatePDF');


Route::group(['middleware' => 'auth'], function () {
    Route::get('/virtual-reality', [PageController::class, 'vr'])->name('virtual-reality');
    Route::get('/rtl', [PageController::class, 'rtl'])->name('rtl');
    Route::get('/profile', [UserProfileController::class, 'show'])->name('profile');
    Route::post('/profile', [UserProfileController::class, 'update'])->name('profile.update');
    Route::get('/profile-static', [PageController::class, 'profile'])->name('profile-static');
    Route::get('/sign-in-static', [PageController::class, 'signin'])->name('sign-in-static');
    Route::get('/sign-up-static', [PageController::class, 'signup'])->name('sign-up-static');
    Route::get('/{page}', [PageController::class, 'index'])->name('page');
    Route::post('logout', [LoginController::class, 'logout'])->name('logout');
});
