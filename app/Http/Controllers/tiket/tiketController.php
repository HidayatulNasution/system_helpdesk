<?php

namespace App\Http\Controllers\tiket;

use App\Models\tiket;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\DB;
use App\Http\Controllers\Controller;
use Illuminate\Support\Facades\File;
use Illuminate\Support\Facades\Response;
use App\Mail\SendEmail;
use Illuminate\Support\Facades\Mail;

class tiketController extends Controller
{
    /**
     * Display a listing of the resource.
     *
     * @return \Illuminate\Http\Response
     */
    public function index()
    {
        if (request()->ajax()) {
            return datatables()->of(tiket::where('user_id', auth()->id())->with('user')->select('*')->where('status', 0))
                ->addColumn('username', function ($row) {
                    return $row->user ? $row->user->username : 'N/A';
                })
                ->addColumn('action', 'tiket.tiket-action')
                ->addColumn('image', 'tiket.image')
                ->rawColumns(['action', 'image'])
                ->addIndexColumn()
                ->make(true);
        }

        // view username dari user yang sedang login
        $username = auth()->user()->username;



        // ambil data persentase berdasarkan tanggal entry
        $reportTiket = tiket::where('user_id', auth()->id())
            ->select(DB::raw('MONTH(created_at) as month'), DB::raw('count(*) as total'))
            ->groupBy('month')
            ->get();

        $totalTiket = $reportTiket->sum('total');

        // array untuk setiap bulan dengan nilai default 0
        $tiketByMonth = array_fill(0, 12, 0);
        foreach ($reportTiket as $tiket) {
            $tiketByMonth[$tiket->month - 1] = $tiket->total;
        }

        // data berdasarkan status 0 (on progress)
        $tiketProgres = tiket::where('user_id', auth()->id())
            ->select(DB::raw('MONTH(created_at) as month'), DB::raw('count(*) as total'))
            ->where('status', 0)
            ->groupBy('month')
            ->get();

        // array untuk data status on progress
        $tiketOnProgres = array_fill(0, 12, 0);
        foreach ($tiketProgres as $tiket) {
            $tiketOnProgres[$tiket->month - 1] = $tiket->total;
        }

        // data berdasarkan status 1 (done)
        $tiketDone = tiket::where('user_id', auth()->id())
            ->select(DB::raw('MONTH(created_at) as month'), DB::raw('count(*) as total'))
            ->where('status', 1)
            ->groupBy('month')
            ->get();

        // array untuk status done
        $tiketOnDone = array_fill(0, 12, 0);
        foreach ($tiketDone  as $tiket) {
            $tiketOnDone[$tiket->month - 1] = $tiket->total;
        }

        return view('tiket.dashboard', compact('tiketByMonth', 'tiketOnProgres', 'tiketOnDone', 'username'));
    }

    public function done()
    {
        if (request()->ajax()) {
            $query = tiket::where('user_id', auth()->id())->with('user')->select('*')->where('status', 1);

            // filter month and year
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
                ->addColumn('action', 'tiket.done-action')
                ->rawColumns(['action', 'image'])
                ->addIndexColumn()
                ->make(true);
        }

        // get the report data as before
        $reportTiket = tiket::where('user_id', auth()->id())
            ->select(DB::raw('MONTH(created_at) as month'), DB::raw('count(*) as total'))
            ->groupBy('month')
            ->get();

        $totalEntries = $reportTiket->sum('total');

        // array data for each month (Jan - Dec) default to 0
        $$dataByMonth = array_fill(0, 12, 0);
        foreach ($reportTiket as $tiket) {
            $dataByMonth[$tiket->month - 1] =  $tiket->total;
        }

        // filter status done
        $statusTiket = tiket::where('user_id', auth()->id())
            ->select(DB::raw('MONTH(created_at) as month'), DB::raw('count(*) as total'))
            ->where('status', 1)
            ->groupBy('month')
            ->get();

        // array data done
        $dataDone = array_fill(0, 12, 0);
        foreach ($statusTiket as $tiket) {
            $dataDone[$tiket->month - 1] = $tiket->total;
        }

        return view('tiket.dashboard', compact('dataByMonth', 'dataDone'));
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
            'image' => 'image|mimes:jpg,png,jpeg,svg,gif|max:2048',
        ]);

        $tiketId = $request->tiket_id;

        $image = $request->hidden_image;

        if ($files = $request->file('image')) {

            // delete old files
            File::delete('public/product/' . $request->hidden_image);

            // insert new file
            $destinationPath = 'public/product/';
            $profileImage = date('YmdHis') . "." . $files->getClientOriginalExtension();
            $files->move($destinationPath, $profileImage);
            $image = "$profileImage";
        }

        $tiket = tiket::find($tiketId) ?? new tiket();
        //$tiket = new tiket();
        $tiket->id = $tiketId;
        $tiket->created_at = $request->created_at;
        $tiket->bidang_system = $request->bidang_system;
        $tiket->kategori = $request->kategori;
        $tiket->status = $request->status;
        $tiket->problem = $request->problem;
        //$tiket->result = $request->result;
        $tiket->prioritas = $request->prioritas;
        $tiket->image = $image;

        // Set user_id untuk tiket baru
        if (!$tiket->exists) {
            $tiket->user_id = auth()->id();
        }

        // save tiket
        $tiket->save();

        // Send email immediately after storing new tiket
        $this->send_email($tiket->id);

        return response()->json(['success' => 'Ticket created and email sent successfully.']);
    }
    /**
     * Display the specified resource.
     *
     * @param  int  $id
     * @return \Illuminate\Http\Response
     */
    public function showDetail($id)
    {
        $tiket = tiket::find($id);
        if (!$tiket) {
            return response()->json(['error' => 'Data Tidak Ditemukan'], 404);
        }

        return response()->json($tiket);
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
        $tiket = tiket::where($where)->first();
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
        $tiket = tiket::where('id', $id)->first(['image']);
        File::delete('public/product/' . $tiket->image);
        $tiket = tiket::where('id', $id)->delete();

        return Response::json($tiket);
    }

    public function send_email($id)
    {
        $tiket = tiket::find($id);
        if (!$tiket) {
            return redirect()->route('home')->with('error', 'Ticket not found.');
        }
        $data = [
            'username' => $tiket->user->username ?? 'N/A',
            'bidang_system' => $tiket->bidang_system,
            'kategori' => $tiket->kategori,
            'prioritas' => $tiket->prioritas == 1 ? 'Urgent' : 'Biasa',
            'problem' => $tiket->problem,
        ];

        $email_target = 'hidayatulnasution9@gmail.com';

        Mail::to($email_target)
            ->cc(['hidayatulafriahman21@gmail.com', 'nasutionhidayatul@gmail.com', 'taurus.silver061@gmail.com'])
            ->send(new SendEmail($data));

        return redirect()->route('home')->with('pesan', 'Email berhasil terkirim.');
    }
}
