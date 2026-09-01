<?php
// ============================================
// dars-jadvali uchun Telegram bildirishnoma API
// Bu faylni cPanel File Manager orqali api.php bilan bir joyga (api.buxpiima.uz) yuklaysiz
// ============================================

header("Access-Control-Allow-Origin: *");
header("Access-Control-Allow-Methods: POST, OPTIONS");
header("Access-Control-Allow-Headers: Content-Type");
header("Content-Type: application/json; charset=utf-8");

if ($_SERVER['REQUEST_METHOD'] === 'OPTIONS') {
    http_response_code(200);
    exit(0);
}

if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
    http_response_code(405);
    echo json_encode(["error" => "Faqat POST so'rov qabul qilinadi"]);
    exit;
}

// ---- BU YERGA @BotFather bergan bot tokenini yozing ----
$BOT_TOKEN = "8953518535:AAEw810GkO6rEk_TagZLOW-RfQ8SY6r4pOo";
// ----------------------------------------------------

$raw = file_get_contents("php://input");
$data = json_decode($raw, true);
$chatId = $data['chatId'] ?? null;
$text = $data['text'] ?? null;

if (!$chatId || !$text) {
    http_response_code(400);
    echo json_encode(["error" => "chatId va text kerak"]);
    exit;
}

$url = "https://api.telegram.org/bot{$BOT_TOKEN}/sendMessage";
$payload = json_encode(["chat_id" => $chatId, "text" => $text]);

$ch = curl_init($url);
curl_setopt($ch, CURLOPT_POST, true);
curl_setopt($ch, CURLOPT_POSTFIELDS, $payload);
curl_setopt($ch, CURLOPT_HTTPHEADER, ["Content-Type: application/json"]);
curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
curl_setopt($ch, CURLOPT_SSL_VERIFYPEER, true);
$result = curl_exec($ch);
$httpCode = curl_getinfo($ch, CURLINFO_HTTP_CODE);
$curlErr = curl_error($ch);
curl_close($ch);

if ($httpCode === 200) {
    echo json_encode(["ok" => true]);
} else {
    http_response_code(500);
    echo json_encode(["ok" => false, "error" => $curlErr ?: $result]);
}
