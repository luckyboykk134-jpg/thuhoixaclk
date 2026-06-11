// api/proxy-gas.js
export default async function handler(req, res) {
  const GOOGLE_SCRIPT_URL = process.env.GOOGLE_SCRIPT_URL;

  if (!GOOGLE_SCRIPT_URL) {
    return res.status(500).json({ error: "Chưa cấu hình biến GOOGLE_SCRIPT_URL trên Vercel" });
  }

  try {
    // --- ĐÂY LÀ ĐOẠN SỬA THEO LƯU Ý ---
    let parsedBody = req.body;
    
    // Nếu React gửi sang dưới dạng chuỗi chữ (String), ta tiến hành biến nó thành Object
    if (typeof req.body === 'string') {
      try {
        parsedBody = JSON.parse(req.body);
      } catch (e) {
        // Nếu không parse được thì giữ nguyên
        parsedBody = req.body;
      }
    }
    // ----------------------------------

    // Chuyển tiếp request đã được xử lý chuẩn sang Google Apps Script
    const response = await fetch(GOOGLE_SCRIPT_URL, {
      method: req.method,
      headers: {
        "Content-Type": "application/json", // Gửi sang Google dưới dạng JSON chuẩn chỉnh
      },
      body: req.method !== 'GET' ? JSON.stringify(parsedBody) : undefined,
    });

    const data = await response.json();
    return res.status(200).json(data);
  } catch (error) {
    return res.status(500).json({ error: "Lỗi kết nối proxy", details: error.message });
  }
}