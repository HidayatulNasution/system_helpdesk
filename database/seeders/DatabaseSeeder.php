<?php

namespace Database\Seeders;

// use Illuminate\Database\Console\Seeds\WithoutModelEvents;

use App\Models\tiket;
use App\Models\User;
use Illuminate\Database\Seeder;
use Illuminate\Support\Facades\DB;

class DatabaseSeeder extends Seeder
{
    /**
     * Seed the application's database.
     *
     * @return void
     */
    public function run()
    {
        DB::table('users')->insert([
            'username' => 'admin',
            'firstname' => 'Admin',
            'lastname' => 'Admin',
            'email' => 'admin@argon.com',
            'password' => bcrypt('123')
        ]);
        tiket::factory(20)->create();
        User::factory(5)->create();
    }
}
