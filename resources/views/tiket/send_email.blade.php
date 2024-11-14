<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta http-equiv="X-UA-Compatible" content="ie=edge">
    <title>New Helpdesk Ticket Notification</title>
</head>

<body>
    <p><strong>SYSTEM HELPDESK</strong></p>

    <p>Terdapat tiket komplain baru dengan detail sebagai berikut:</p>

    <ul>
        <li><strong>User:</strong> {{ $data['username'] }}</li>
        <li><strong>Sistem:</strong> {{ $data['bidang_system'] }}</li>
        <li><strong>Menu:</strong> {{ $data['kategori'] }}</li>
        <li><strong>Sifat:</strong> {{ $data['prioritas'] }}</li>
        <li><strong>Problem:</strong> {{ $data['problem'] }}</li>
    </ul>

    <p>Demikian disampaikan, atas perhatian dan kerjasamanya yang baik saya ucapkan terima kasih.</p>

    <p>Best Regards,</p>
    <strong class="uppercase">{{ $data['username'] }}</strong>
    <br>

    <i><strong>Ini adalah pesan otomatis, mohon untuk tidak membalas pesan ini.</strong></i>
</body>

</html>
