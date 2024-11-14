<!DOCTYPE html>
<html>

<head>
    <title>Report Tiket</title>
    <!-- Argon Bootstrap CSS -->
    <link href="https://cdnjs.cloudflare.com/ajax/libs/argon-design-system/1.2.0/css/argon.min.css" rel="stylesheet">
    <!-- Bootstrap Icons (optional) -->
    <link href="https://cdnjs.cloudflare.com/ajax/libs/bootstrap-icons/1.8.1/font/bootstrap-icons.min.css"
        rel="stylesheet">
    <style>
        .table th,
        .table td {
            vertical-align: middle;
            text-align: center;
        }

        h1 {
            font-weight: bold;
        }

        .table thead {
            background-color: #f6f9fc;
        }
    </style>
    <style type="text/css">
        @page {
            margin-top: 0.8cm;
            margin-left: 1cm;
            margin-right: 1cm;
            margin-bottom: 0.8cm;
        }

        .font {
            font-family: verdana, arial, sans-serif;
            font-size: 14px;
        }

        .fontheader {
            font-family: verdana, arial, sans-serif;
            font-size: 13px;
            color: #333333;
            border-width: 1px;
            border-color: #666666;
            border-collapse: collapse;
        }

        table.noborder {
            font-family: verdana, arial, sans-serif;
        }

        table.noborder th {
            font-size: 12px;
            padding: 1px;
            border-color: #666666;
        }

        table.noborder td {
            font-size: 12px;
            padding: 1px;
            border-color: #666666;
            background-color: #ffffff;
        }

        table.gridtable {
            font-family: verdana, arial, sans-serif;
            font-size: 11px;
            color: #333333;
            border-width: 1px;
            border-color: #666666;
            border-collapse: collapse;
        }

        table.gridtable th {
            border-width: 1px;
            padding: 5px;
            border-style: solid;
            border-color: #666666;
            background-color: #f2f2f2;
        }

        table.gridtable td {
            border-width: 1px;
            padding: 5px;
            border-style: solid;
            border-color: #666666;
            background-color: #ffffff;
        }

        table.gridtable td.zero {
            border-width: 1px;
            padding: 5px;
            border-color: #666666;
            background-color: #ffffff;
        }

        table.gridtable td.cols {
            border-width: 1px;
            padding: 5px;
            border-style: solid;
            border-color: #666666;
            background-color: #ffffff;
        }
    </style>
</head>

<body class="bg-light">
    <div class="container my-5">
        <h1 class="text-center mb-4">{{ $title }}</h1>

        <table class="noborder" width='100%' style='page-break-inside:auto;'>
            <tr>
                <td align='left' width='15%'>Periode</td>
                <td align='left' width='2%'>:</td>
                <td align='left' width='78%'>
                    {{ \Carbon\Carbon::parse($tikets->first()->created_at)->format('d F Y') }}
                    {{-- -
                    {{ \Carbon\Carbon::parse($tikets->last()->updated_at)->format('d F Y') }} --}}
                </td>
            </tr>
        </table>

        <br>

        <table class="gridtable" width='100%'>
            <tr>
                <th align='left' colspan='10'># Detail</th>
            </tr>
            <tr>
                <th width='5%' align='center'>No</th>
                <th width='10%' align='center'>User</th>
                <th width='15%' align='center'>Created At</th>
                <th width='31%' align='center'>Bidang System</th>
                <th width='31%' align='center'>Kategori</th>
                <th width='13%' align='center'>Prioritas</th>
                <th width='35%' align='center'>Problem</th>
                <th width='35%' align='center'>Result</th>
                <th width='10%' align='center'>Status</th>
                <th width='15%' align='center'>Tanggal Selesai</th>
            </tr>

            <tbody>
                @php $loopIndex = 1; @endphp

                @foreach ($tikets as $tiket)
                    <tr>
                        <td width='5%' align='center'>{{ $loopIndex++ }}</td>
                        <td width='10%' align='center'>{{ $tiket->user->username }}</td>
                        <td width='15%' align='center'>
                            {{ \Carbon\Carbon::parse($tiket->created_at)->format('d/m/Y H:i') }}</td>
                        <td width='31%' align='center'>{{ $tiket->bidang_system }}</td>
                        <td width='31%' align='center'>{{ $tiket->kategori }}</td>
                        <td width='13%' align='center'
                            class="{{ $tiket->prioritas == 1 ? 'text-success' : 'text-warning' }}">
                            {{ $tiket->prioritas == 1 ? 'URGENT' : 'BIASA' }}
                        </td>
                        <td width='35%' align='center'>{{ $tiket->problem }}</td>
                        <td width='35%' align='center'>{{ $tiket->result }}</td>
                        <td width='10%' align='center'
                            class="{{ $tiket->status == 1 ? 'text-danger' : 'text-success' }}">
                            {{ $tiket->status == 1 ? 'DONE' : 'On Progress' }}
                        </td>
                        <td width='15%' align='center'>
                            {{ \Carbon\Carbon::parse($tiket->updated_at)->format('d/m/Y H:i') }}</td>
                    </tr>
                @endforeach
            </tbody>
        </table>
    </div>

    <br><br>
    <p class='footer'>
        <strong>Time Printed: {{ $date }}</strong>
    </p>

    <!-- Argon Bootstrap JS and dependencies (if needed) -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/argon-design-system/1.2.0/js/argon.min.js"></script>
</body>

</html>
