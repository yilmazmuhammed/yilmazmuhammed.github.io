// api/davetli-getir.js

export default function handler(req, res) {
    // Bu liste SADECE sunucuda kalacak, tarayıcıya gitmeyecek
    const davetliler = {
        "44njSN": "Ahmet Said Şenol",
        "xYz123": "Örnek İsim 2"
        // Python scriptiniz bu kısmı güncelleyebilir...
    };

    // Kullanıcının URL'den gönderdiği token'ı alıyoruz (?t=44njSN)
    const { t } = req.query; 

    // Eğer token yoksa veya listede eşleşmiyorsa 404 hatası dön
    if (!t || !davetliler[t]) {
        return res.status(404).json({ mesaj: "Davetli bulunamadı veya geçersiz token." });
    }

    // Eşleşme varsa sadece o kişinin ismini JSON olarak dön
    return res.status(200).json({ isim: davetliler[t] });
}
