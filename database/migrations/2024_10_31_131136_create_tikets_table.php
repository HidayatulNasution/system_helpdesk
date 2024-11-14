<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

return new class extends Migration
{
    /**
     * Run the migrations.
     *
     * @return void
     */
    public function up()
    {
        Schema::create('tikets', function (Blueprint $table) {
            $table->id();
            $table->string('bidang_system');
            $table->string('kategori');
            $table->boolean('status')->default(false);
            $table->text('problem')->nullable();
            $table->text('result')->nullable();
            $table->boolean('prioritas')->default(false);
            $table->string('image')->nullable();
            $table->foreignId('user_id')->nullable()->constrained('users')->onDelete('cascade'); // Menambahkan user_id dengan foreign key
            $table->timestamps();
        });
    }

    /**
     * Reverse the migrations.
     *
     * @return void
     */
    public function down()
    {
        Schema::dropIfExists('tikets');
    }
};
