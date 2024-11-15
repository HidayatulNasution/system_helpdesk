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
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Font;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;

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

        // If no data is found, return with an error message
        if ($tikets->isEmpty()) {
            return redirect()->back()->with('error', 'No Data Available');
        }

        $data = [
            'title' => 'Report Tiket',
            'date' => date('m/d/Y'),
            'tikets' => $tikets
        ];

        $pdf = PDF::loadView('admin.tiketsPdf', $data);
        return $pdf->download('Report_Tiket_' . now()->format('d-m-Y_His') . '.pdf');
    }

    public function exportTiketsToExcel()
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

        // Jika data tidak ditemukan, kirim respons SweetAlert
        if ($tikets->isEmpty()) {
            return response()->json(['status' => 'error', 'message' => 'Download Summary Data']);
        }

        // Tentukan periode berdasarkan created_at
        $startDate = $tikets->min('created_at');
        $endDate = $tikets->max('created_at');
        $period = 'Periode: ' . $startDate->format('d/m/Y') . ' - ' . $endDate->format('d/m/Y');

        // Membuat objek Spreadsheet baru
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();

        // Judul di tengah
        $sheet->setCellValue('A1', 'Report Tiket');
        $sheet->mergeCells('A1:H1');
        $sheet->getStyle('A1')->getFont()->setBold(true)->setSize(20);
        $sheet->getStyle('A1')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

        // Tambahkan periode di bawah judul
        $sheet->setCellValue('A2', $period);
        $sheet->mergeCells('A2:H2');
        $sheet->getStyle('A2')->getFont()->setItalic(true)->setSize(12);
        $sheet->getStyle('A2')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);

        // Gaya untuk header
        $sheet->setCellValue('A4', 'No');
        $sheet->setCellValue('B4', 'Username');
        $sheet->setCellValue('C4', 'Bidang System');
        $sheet->setCellValue('D4', 'Kategori');
        $sheet->setCellValue('E4', 'Status');
        $sheet->setCellValue('F4', 'Problem');
        $sheet->setCellValue('G4', 'Result');
        $sheet->setCellValue('H4', 'Prioritas');

        $headerStyle = [
            'font' => [
                'bold' => true,
                'size' => 12,
                'color' => ['rgb' => 'FFFFFF']
            ],
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_CENTER,
                'vertical' => Alignment::VERTICAL_CENTER,
            ],
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => ['rgb' => '4CAF50']
            ],
            'borders' => [
                'allBorders' => [
                    'borderStyle' => Border::BORDER_THIN,
                    'color' => ['rgb' => '000000'],
                ]
            ],
        ];

        // Terapkan gaya ke header
        $sheet->getStyle('A4:H4')->applyFromArray($headerStyle);

        // Gaya untuk data rata kiri
        $leftAlignedStyle = [
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_LEFT,
                'vertical' => Alignment::VERTICAL_CENTER,
            ],
            'borders' => [
                'allBorders' => [
                    'borderStyle' => Border::BORDER_THIN,
                    'color' => ['rgb' => '000000'],
                ]
            ],
        ];

        // Gaya untuk data rata tengah
        $centerAlignedStyle = [
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_CENTER,
                'vertical' => Alignment::VERTICAL_CENTER,
            ],
            'borders' => [
                'allBorders' => [
                    'borderStyle' => Border::BORDER_THIN,
                    'color' => ['rgb' => '000000'],
                ]
            ],
        ];

        // Isi data ke dalam tabel Excel
        $row = 5;
        $nomorUrut = 1; // Menambahkan nomor urut
        foreach ($tikets as $tiket) {
            $statusText = $tiket->status == 1 ? 'DONE' : 'On Progress';
            $prioritasText = $tiket->prioritas == 1 ? 'URGENT' : 'BIASA';
            $username = $tiket->user ? $tiket->user->username : 'N/A'; // Cek jika username tersedia

            $sheet->setCellValue('A' . $row, $nomorUrut);
            $sheet->setCellValue('B' . $row, $username);
            $sheet->setCellValue('C' . $row, $tiket->bidang_system);
            $sheet->setCellValue('D' . $row, $tiket->kategori);
            $sheet->setCellValue('E' . $row, $statusText);
            $sheet->setCellValue('F' . $row, $tiket->problem);
            $sheet->setCellValue('G' . $row, $tiket->result);
            $sheet->setCellValue('H' . $row, $prioritasText);

            // Terapkan gaya pada kolom
            $sheet->getStyle('B' . $row)->applyFromArray($leftAlignedStyle);
            $sheet->getStyle('C' . $row)->applyFromArray($leftAlignedStyle);
            $sheet->getStyle('D' . $row)->applyFromArray($leftAlignedStyle);
            $sheet->getStyle('F' . $row)->applyFromArray($leftAlignedStyle);
            $sheet->getStyle('G' . $row)->applyFromArray($leftAlignedStyle);
            $sheet->getStyle('A' . $row)->applyFromArray($centerAlignedStyle);
            $sheet->getStyle('E' . $row)->applyFromArray($centerAlignedStyle);
            $sheet->getStyle('H' . $row)->applyFromArray($centerAlignedStyle);

            $row++;
            $nomorUrut++;
        }

        // Set lebar kolom agar lebih rapi
        foreach (range('A', 'H') as $columnID) {
            $sheet->getColumnDimension($columnID)->setAutoSize(true);
        }

        // Simpan file Excel
        $writer = new Xlsx($spreadsheet);
        $fileName = 'report_tiket_' . now()->format('d-m-Y_His') . '.xlsx';
        $path = storage_path('app/public/' . $fileName);
        $writer->save($path);

        // Memberikan respons untuk mengunduh file
        return response()->download($path);
    }
}
