<?php
namespace App\Http\Controllers;
use Illuminate\Http\Request;
use App\Http\Models\Transaksi;
use Box\Spout\Reader\ReaderFactory;
use Box\Spout\Writer\WriterFactory;
use Box\Spout\Common\Type;

class ExcelController extends Controller
{
    public function index()
    {
        echo "export excel using spout";
    }
    
    public function upload(Request $request)
    {
        if ($request->hasFile('file')) {
            $file = $request->file('file');
            $reader = ReaderFactory::create(Type::XLSX); 
            $reader->open($file);
            
            foreach ($reader->getSheetIterator() as $sheet) {
                if ($sheet->getName() === 'Orders') {
                    $this->readOrderSheet($sheet);
                }
            }
            $reader->close();
        }
    }
    
    public function readOrderSheet($sheet)
    {
        foreach ($sheet->getRowIterator() as $idx => $row) {
            if ($idx>1) { 
                $data = [
                'transaksi_id' => $row[0],
                'transaksi_serial_number' => $row[1],
                'transaksi_tanggal_transaksi_kartu' => $row[2],
                'transaksi_transaksi_istransit' => $row[3],
                'transaksi_jns_pengguna' => $row[4],
                'transaksi_jns_kartu' => $row[5],
                'transaksi_nominal' => $row[6],
                'transaksi_kode_channel' => $row[7],
                'transaksi_moda' => $row[8],
                'transaksi_lokasi' => $row[9],
                'transaksi_lokasi_koordinat' => $row[10],
                'transaksi_status' => $row[11],
                'transaksi_issuer' => $row[12],
                'transaksi_counter' => $row[13],
                'transaksi_pengiriman_timestamp' => $row[14],
                'transaksi_sent_status' => $row[15],
                'transaksi_approved_status' => $row[16],
                'transaksi_unapproved_status' => $row[17],
                'transaksi_header' => $row[18]
                ];
                $transaksi = new Transaksi();
                $transaksi->fill($data);
                $transaksi->save(); 
            }
        }
    }

    public function exportExcel()
    {
        $title = [
            'Transaksi Id', 
            'Serial Number', 
            'Tanggal Transaksi', 
            'Nominal', 
            'Kartu'
        ];

        $fileName = 'Export Excel.xlsx';
        $writer = WriterFactory::create(Type::XLSX); 

        // data
        $transaksi = Transaksi::select(
            'transaksi_id', 
            'transaksi_serial_number', 
            'transaksi_tanggal_transaksi_kartu', 
            'transaksi_nominal', 
            'transaksi_jns_kartu'
        )->get();
        // **
        
        $writer->openToBrowser($fileName);
        $writer->addRow($title); 

        foreach ($transaksi as $idx => $data)
            $writer->addRow($data->toArray());

        $writer->close();
    }
}