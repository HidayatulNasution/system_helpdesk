<?php

namespace App\Http\Controllers\admin;

use App\Models\User;
use App\Models\admin;
use Illuminate\Http\Request;
use Barryvdh\DomPDF\Facade\Pdf;
use Illuminate\Support\Facades\DB;
use App\Http\Controllers\Controller;
use Illuminate\Support\Facades\File;
use Illuminate\Support\Facades\Response;

class adminController extends Controller
{
    /**
     * Display a listing of the resource.
     *
     * @return \Illuminate\Http\Response
     */
    public function index()
    {
        if (request()->ajax()) {
            return datatables()->of(admin::with('user')->select('*')->where('status', 0))
                ->addColumn('username', function ($row) {
                    return $row->user ? $row->user->username : 'N/A';
                })
                ->addColumn('action', 'admin.admin-action')
                ->addColumn('image', 'tiket.image')
                ->rawColumns(['action', 'image'])
                ->addIndexColumn()
                ->make(true);
        }

        // data persentase berdasarkan tanggal entry
        $reportData = admin::select(DB::raw('MONTH(created_at) as month'), DB::raw('count(*) as total'))
            ->groupBy('month')
            ->get();

        $totalEntries = $reportData->sum('total');

        // array untuk setiap bulan
        $dataByMonth = array_fill(0, 12, 0);
        foreach ($reportData as $data) {
            $dataByMonth[$data->month - 1] = $data->total;
        }

        // data berdasarkan status on progress
        $statusProgres = admin::select(DB::raw('MONTH(created_at) as month'), DB::raw('count(*) as total'))
            ->where('status', 0)
            ->groupBy('month')
            ->get();

        // array status on progress
        $dataProgres = array_fill(0, 12, 0);
        foreach ($statusProgres as $data) {
            $dataProgres[$data->month - 1] = $data->total;
        }

        // data berdasarkan status done
        $statusData = admin::select(DB::raw('MONTH(created_at) as month'), DB::raw('count(*) as total'))
            ->where('status', 1)
            ->groupBy('month')
            ->get();

        // array status done
        $dataByStatus = array_fill(0, 12, 0);
        foreach ($statusData as $data) {
            $dataByStatus[$data->month - 1] = $data->total;
        }

        // Mengambil semua tiket beserta data pengguna yang terkait
        $tikets = admin::with('user')->get(); // Menggunakan eager loading untuk relasi

        return view('admin.index', compact('dataByMonth', 'dataProgres', 'dataByStatus', 'tikets'));
    }

    public function done()
    {
        if (request()->ajax()) {
            $query = admin::with('user')->select('*')->where('status', 1);

            // filter month and year if provided
            $month = request()->get('month');
            $year = request()->get('year');
            if ($month) {
                $query->whereMonth('created_at', $month);
            }
            if ($year) {
                $query->whereYear('created_at', $year);
            }

            return datatables()->of($query)
                ->addColumn('username', function ($row) {
                    return $row->user ? $row->user->username : 'N/A';
                })
                ->addColumn('action', 'admin.dones-action')
                ->rawColumns(['action', 'image'])
                ->addIndexColumn()
                ->make(true);
        }

        // report data
        $reportData = admin::select(DB::raw('MONTH(created_at) as month'), DB::raw('count(*) as total'))
            ->groupBy('month')
            ->get();

        $totalEntries = $reportData->sum('total');

        // array data for each month 
        $dataByMonth = array_fill(0, 12, 0);
        foreach ($reportData as $data) {
            $dataByMonth[$data->month - 1] = $data->total;
        }

        // data based on status done
        $statusData = admin::select(DB::raw('MONTH(created_at) as month'), DB::raw('count(*) as total'))
            ->where('status', 1)
            ->groupBy('month')
            ->get();

        // array data status done
        $dataDone =  array_fill(0, 12, 0);
        foreach ($statusData as $data) {
            $dataDone[$data->month - 1] = $data->total;
        }

        // Mengambil semua tiket beserta data pengguna yang terkait
        $tikets = admin::with('user')->get(); // Menggunakan eager loading untuk relasi

        return view('admin.index', compact('dataByMonth', 'dataDone', 'tikets'));
    }
    /**
     * Store a newly created resource in storage.
     *
     * @param  \Illuminate\Http\Request  $request
     * @return \Illuminate\Http\Response
     */
    public function store(Request $request)
    {
        request()->validate([
            'image' => 'image|mimes:jpeg,png,gif,svg|max:2048',
        ]);

        $tiketId = $request->tiket_id;

        $image = $request->hidden_image;

        if ($files = $request->file('image')) {

            // delete file
            File::delete('public/product/' . $request->hidden_image);

            // insert new file
            $destinationPath = 'public/product/';
            $profileImage = date('YmdHis') . "." . $files->getClientOriginalExtension();
            $files->move($destinationPath, $profileImage);
            $image = "$profileImage";
        }

        $tiket = admin::find($tiketId) ?? new admin();
        $tiket->id = $tiketId;
        $tiket->created_at = $request->created_at;
        $tiket->bidang_system = $request->bidang_system;
        $tiket->kategori = $request->kategori;
        $tiket->status = $request->status;
        $tiket->problem = $request->problem;
        $tiket->result = $request->result;
        $tiket->prioritas = $request->prioritas;
        $tiket->image = $image;

        // save tiket
        $tiket->save();

        return Response::json($tiket);
    }
    /**
     * Show the form for editing the specified resource.
     *
     * @param  int  $id
     * @return \Illuminate\Http\Response
     */
    public function edit($id)
    {
        $where = array('id' => $id);
        $tiket =  admin::where($where)->first();

        return Response::json($tiket);
    }
    /**
     * Remove the specified resource from storage.
     *
     * @param  int  $id
     * @return \Illuminate\Http\Response
     */
    public function destroy($id)
    {
        $data = admin::where('id', $id)->first(['image']);
        File::delete('public/product/' . $data->image);
        $tiket = admin::where('id', $id)->delete();

        return Response::json($tiket);
    }

    public function showDetail($id)
    {
        $tiket = admin::find($id);
        if (!$tiket) {
            return response()->json(['error' => 'Data Tidak Ditemukan'], 404);
        }

        return response()->json($tiket);
    }

    public function generatePDF()
    {
        $tikets = admin::with('user')->select('*')->where('status', 1);
        $month = request()->get('month');
        $year = request()->get('year');

        // Filter berdasarkan bulan jika ada
        if ($month) {
            $tikets->whereMonth('created_at', $month);
        }

        // Filter berdasarkan tahun jika ada
        if ($year) {
            $tikets->whereYear('created_at', $year);
        }

        $tikets = $tikets->get();

        $data = [
            'title' => 'Report Tiket',
            'date' => date('m/d/Y'),
            'tikets' => $tikets
        ];

        $pdf = PDF::loadView('admin.tiketsPdf', $data);
        return $pdf->download('Report_Tiket.pdf');
    }
}
