<?php
// ============================================
// dars-jadvali uchun oddiy PHP + MySQL API
// Bu faylni cPanel File Manager orqali serveringizga yuklaysiz
// ============================================

header("Access-Control-Allow-Origin: *");
header("Access-Control-Allow-Methods: GET, POST, OPTIONS");
header("Access-Control-Allow-Headers: Content-Type");
header("Content-Type: application/json; charset=utf-8");

if ($_SERVER['REQUEST_METHOD'] === 'OPTIONS') {
    http_response_code(200);
    exit(0);
}

// ---- BU YERGA cPanel'da yaratgan ma'lumotlaringizni kiriting ----
$DB_HOST = "localhost";
$DB_NAME = "xorazmi1_darsjadvali";   // cPanel MySQL Databases'da yaratgan baza nomi
$DB_USER = "xorazmi1_Farid";        // cPanel MySQL Databases'da yaratgan foydalanuvchi
$DB_PASS = "@raxmonov5045"; // shu foydalanuvchining paroli
// -------------------------------------------------------------

$mysqli = new mysqli($DB_HOST, $DB_USER, $DB_PASS, $DB_NAME);
if ($mysqli->connect_error) {
    http_response_code(500);
    echo json_encode(["error" => "Bazaga ulanib bo'lmadi: " . $mysqli->connect_error]);
    exit;
}
$mysqli->set_charset("utf8mb4");

$key = null;
if (isset($_GET['key'])) {
    $key = $_GET['key'];
} elseif ($_SERVER['REQUEST_METHOD'] === 'POST') {
    // POST so'rovda ham ?key= orqali keladi (frontend shunday yuboradi)
    $key = $_GET['key'] ?? null;
}

if (!$key) {
    http_response_code(400);
    echo json_encode(["error" => "key parametri kerak"]);
    exit;
}

if ($_SERVER['REQUEST_METHOD'] === 'GET') {
    $stmt = $mysqli->prepare("SELECT value FROM kv_storage WHERE storage_key = ? LIMIT 1");
    $stmt->bind_param("s", $key);
    $stmt->execute();
    $result = $stmt->get_result();
    if ($row = $result->fetch_assoc()) {
        echo json_encode(["value" => $row['value']]);
    } else {
        echo json_encode(["value" => null]);
    }
    $stmt->close();
} elseif ($_SERVER['REQUEST_METHOD'] === 'POST') {
    $raw = file_get_contents("php://input");
    $data = json_decode($raw, true);
    $value = $data['value'] ?? null;
    if ($value === null) {
        http_response_code(400);
        echo json_encode(["error" => "value kerak"]);
        exit;
    }
    $stmt = $mysqli->prepare(
        "INSERT INTO kv_storage (storage_key, value) VALUES (?, ?)
         ON DUPLICATE KEY UPDATE value = VALUES(value)"
    );
    $stmt->bind_param("ss", $key, $value);
    $stmt->execute();
    echo json_encode(["ok" => true]);
    $stmt->close();
} else {
    http_response_code(405);
    echo json_encode(["error" => "Ruxsat etilmagan usul"]);
}

$mysqli->close();
