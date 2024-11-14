<?php

namespace App\Http\Controllers;

//use Response;
use DataTables;
use App\Models\Product;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\DB;
use Illuminate\Support\Facades\File;
use Illuminate\Support\Facades\Response;

class ProductController extends Controller
{
    public function index()
    {
        if (request()->ajax()) {
            return datatables()->of(Product::select('*')->where('status', 0))
                ->addColumn('action', 'pages.product-action')
                ->addColumn('image', 'product.image')
                ->rawColumns(['action', 'image'])
                ->addIndexColumn()
                ->make(true);
        }

        // Ambil data persentase berdasarkan tanggal entry
        $reportData = Product::select(DB::raw('MONTH(created_at) as month'), DB::raw('count(*) as total'))
            ->groupBy('month')
            ->get();

        $totalEntries = $reportData->sum('total');

        // Buat array data untuk setiap bulan (Jan - Dec) dengan nilai default 0
        $dataByMonth = array_fill(0, 12, 0);
        foreach ($reportData as $data) {
            $dataByMonth[$data->month - 1] = $data->total;
        }

        // Tambahkan data berdasarkan status on progress
        $statusProgres = Product::select(DB::raw('MONTH(created_at) as month'), DB::raw('count(*) as total'))
            ->where('status', 0) // Anggap 1 adalah 'DONE' dan 0 adalah 'On Progress'
            ->groupBy('month')
            ->get();

        // Buat array data untuk status 'DONE' per bulan
        $dataProgres = array_fill(0, 12, 0);
        foreach ($statusProgres as $data) {
            $dataProgres[$data->month - 1] = $data->total;
        }

        // Tambahkan data berdasarkan status
        $statusData = Product::select(DB::raw('MONTH(created_at) as month'), DB::raw('count(*) as total'))
            ->where('status', 1) // Anggap 1 adalah 'DONE' dan 0 adalah 'On Progress'
            ->groupBy('month')
            ->get();

        // Buat array data untuk status 'DONE' per bulan
        $dataByStatus = array_fill(0, 12, 0);
        foreach ($statusData as $data) {
            $dataByStatus[$data->month - 1] = $data->total;
        }

        // return view('product.home');
        return view('pages.dashboard', compact('dataByMonth', 'dataProgres', 'dataByStatus'));
    }

    public function done()
    {
        if (request()->ajax()) {
            $query = Product::select('*')->where('status', 1);

            // Apply month and year filters if provided
            $month = request()->get('month');
            $year = request()->get('year');
            if ($month) {
                $query->whereMonth('created_at', $month);
            }
            if ($year) {
                $query->whereYear('created_at', $year);
            }

            return datatables()->of($query)
                ->addColumn('action', 'pages.done-action')
                ->rawColumns(['action', 'image'])
                ->addIndexColumn()
                ->make(true);
        }

        // Get the report data as before
        $reportData = Product::select(DB::raw('MONTH(created_at) as month'), DB::raw('count(*) as total'))
            ->groupBy('month')
            ->get();
        $totalEntries = $reportData->sum('total');

        // Array data for each month (Jan - Dec) default to 0
        $dataByMonth = array_fill(0, 12, 0);
        foreach ($reportData as $data) {
            $dataByMonth[$data->month - 1] = $data->total;
        }

        // Add data based on status
        $statusData = Product::select(DB::raw('MONTH(created_at) as month'), DB::raw('count(*) as total'))
            ->where('status', 1)
            ->groupBy('month')
            ->get();

        // Array data for status 'DONE' per month
        $dataDone = array_fill(0, 12, 0);
        foreach ($statusData as $data) {
            $dataDone[$data->month - 1] =  $data->total;
        }

        return view('pages.dashboard', compact('dataByMonth', 'dataDone'));
    }

    public function store(Request $request)
    {
        request()->validate([
            'image' => 'image|mimes:jpeg,png,jpg,gif,svg|max:2048',
        ]);
        $productId = $request->product_id;

        $image = $request->hidden_image;

        if ($files = $request->file('image')) {

            //delete old file
            File::delete('public/product/' . $request->hidden_image);

            //insert new file
            $destinationPath = 'public/product/'; // upload path
            $profileImage = date('YmdHis') . "." . $files->getClientOriginalExtension();
            $files->move($destinationPath, $profileImage);
            $image = "$profileImage";
        }

        $product = Product::find($productId) ?? new Product();
        // Set the individual attributes
        $product->id = $productId;
        $product->title = $request->title;
        $product->category = $request->category;
        $product->price = $request->price;
        $product->status = $request->status;
        $product->image = $image;


        // Save the product
        $product->save();

        return Response::json($product);
    }

    public function edit($id)
    {
        $where = array('id' => $id);
        $product  = Product::where($where)->first();

        return Response::json($product);
    }

    public function destroy($id)
    {
        $data = Product::where('id', $id)->first(['image']);
        File::delete('public/product/' . $data->image);
        $product = Product::where('id', $id)->delete();

        return Response::json($product);
    }

    public function showDetail($id)
    {
        $product = Product::find($id);
        if (!$product) {
            return response()->json(['error' => 'Data Tidak Ditemukan'], 404);
        }
        return response()->json($product);
    }
}
