@extends('layouts.app')

@section('content')
    @include('layouts.navbars.auth.topnav', ['title' => 'User Management'])
    @include('pages.table-user')
    @include('pages.modal-user')
@endsection
