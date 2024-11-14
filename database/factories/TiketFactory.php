<?php

namespace Database\Factories;

use App\Models\User;
use Illuminate\Database\Eloquent\Factories\Factory;

/**
 * @extends \Illuminate\Database\Eloquent\Factories\Factory<\App\Models\tiket>
 */
class TiketFactory extends Factory
{
    /**
     * Define the model's default state.
     *
     * @return array<string, mixed>
     */
    public function definition()
    {
        return [
            'bidang_system' => fake()->sentence(),
            'kategori' => fake()->sentence(),
            'status' => fake()->randomElement([0, 1]),
            'problem' => fake()->sentence(),
            //'result' => fake()->sentence(),
            'user_id' => User::factory(),
            'prioritas' => fake()->randomElement([0, 1])
        ];
    }
}
