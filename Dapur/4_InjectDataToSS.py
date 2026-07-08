Otak yang akan memindai data di spreadsheets orderan berdasarkan nomor pelanggan dan menggunakan fungsi perbatch untuk menghemat API
Menyisipkan total piutang, tidak termasuk fraud, menyisipkan keterangan owing dan JT jika ada. harapan isi:
SL-1151	EKA SAKTI	PCMO (dari excel AR)
07/07/2026		(ambil tanggal masuk orderan)
		
Piutang	 30.002.000 	(total yang sama dengan kalkulasi piutang, tidak termasuk fraud)
Inv	 4 	(total faktur)
Plafon	 13.800.000 	(ambil dari excel hasil unduhan spreadsheets untuk AVG AR)
Rata" Bayar	 27.636.400 	(ambil dari excel hasil unduhan spreadsheets untuk AVG AR)
Rata" History Bayar	99	Hari (ambil dari excel hasil unduhan spreadsheets untuk AVG AR)

260600522	 5.470.400 	 33 	HR (JT 08/07/26) (dari excel AR + hasl helper giro DD/MM/YY)
260605984	 7.654.100 	 10 	HR (contoh biasa tanpa giro)
