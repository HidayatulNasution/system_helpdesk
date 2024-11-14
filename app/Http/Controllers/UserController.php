<?php


namespace App\Http\Controllers;


use App\Models\User;
use Illuminate\Http\Request;
use Illuminate\Http\Response;
use Illuminate\Support\Facades\DB;
use App\Http\Controllers\Controller;
use Illuminate\Support\Facades\File;
use Illuminate\Support\Facades\Hash;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class UserController extends Controller
{
    public function index()
    {
        if (request()->ajax()) {
            return datatables()->of(User::select('*'))
                ->addColumn('action', 'pages.user-action')
                ->addIndexColumn()
                ->make(true);
        }

        return view('pages.manage-user');
    }

    /**
     * Store a newly created resource in storage.
     *
     * @param  \Illuminate\Http\Request  $request
     * @return \Illuminate\Http\Response
     */
    public function store()
    {
        $attributes = request()->validate([
            'username' => 'max:255',
            'email' => 'max:255',
            'password' => 'max:255',
            'status' => 'required'
        ]);

        $user = User::create($attributes);
        auth()->login($user);

        return redirect('/user');
    }

    /**
     * Display the specified resource.
     *
     * @param  int  $id
     * @return \Illuminate\Http\Response
     */
    public function showDetail($id)
    {
        $user = User::find($id);
        if (!$user) {
            return response()->json(['error' => 'Data Tidak Ditemukan'], 404);
        }
        return response()->json($user);
    }

    /**
     * Show the form for editing the specified resource.
     *
     * @param  int  $id
     * @return \Illuminate\Http\Response
     */
    public function edit($id)
    {
        $user = User::find($id);
        if (!$user) {
            return response()->json(['error' => 'Data Tidak Ditemukan'], 404);
        }
        return response()->json($user);
    }

    /**
     * Remove the specified resource from storage.
     *
     * @param  int  $id
     * @return \Illuminate\Http\Response
     */
    public function destroy($id)
    {
        $user  = User::find($id);

        $user->delete();

        return response()->json(['success' => true]);
    }
}
