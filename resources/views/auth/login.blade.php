@extends('layouts.app')

@section('content')
    {{-- <div class="container position-sticky z-index-sticky top-0">
        <div class="row">
            <div class="col-12">
                @include('layouts.navbars.guest.navbar')
            </div>
        </div>
    </div> --}}
    <main class="main-content">
        <section>
            <div class="page-header min-vh-100">
                <div class="container">
                    <div class="row justify-content-start">
                        <div class="col-xl-4 col-lg-5 col-md-7 d-flex flex-column mx-lg-0 mx-auto">
                            <div class="card card-plain">
                                <div class="card-header pb-0 text-start">
                                    <h4 class="font-weight-bolder">Sign In</h4>
                                    <p class="mb-0">Enter your username and password to sign in</p>
                                </div>
                                <div class="card-body">
                                    <form role="form" method="POST" action="{{ route('login.perform') }}">
                                        @csrf
                                        @method('post')
                                        <div class="flex flex-col mb-3">
                                            <input type="username" name="username" class="form-control form-control-lg"
                                                aria-label="username" value="" placeholder="username">
                                            @error('username')
                                                <p class="text-danger text-xs pt-1"> {{ $message }} </p>
                                            @enderror
                                        </div>

                                        <div class="flex flex-col mb-3">
                                            <input type="password" name="password" class="form-control form-control-lg"
                                                aria-label="Password" value="" placeholder="password">
                                            @error('password')
                                                <p class="text-danger text-xs pt-1"> {{ $message }} </p>
                                            @enderror
                                        </div>
                                        <div class="form-check form-switch">
                                            <input class="form-check-input" name="remember" type="checkbox" id="rememberMe">
                                            <label class="form-check-label" for="rememberMe">Remember me</label>
                                        </div>
                                        <div class="text-center">
                                            <button type="submit"
                                                class="btn btn-lg btn-primary btn-lg w-100 mt-4 mb-0">Sign in</button>
                                        </div>
                                    </form>
                                </div>
                            </div>
                            @push('js')
                                <!-- SweetAlert Success Notification -->
                                @if (session('login_success'))
                                    <script src="https://cdn.jsdelivr.net/npm/sweetalert2@11"></script>
                                    <script>
                                        Swal.fire({
                                            title: "Success!",
                                            text: "{{ session('login_success') }}",
                                            icon: "success",
                                            timer: 3000,
                                            showConfirmButton: false
                                        });
                                    </script>
                                @endif
                            @endpush

                        </div>
                        <div class="col-md-6 col-lg-6 position-absolute  top-0 end-10">
                            <img src="/img/login2.jpg" alt="pdf">
                        </div>
                    </div>

                </div>
            </div>
            </div>
        </section>
    </main>
@endsection
