import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Scanner;

import com.artofsolving.jodconverter.DefaultDocumentFormatRegistry;
import com.artofsolving.jodconverter.DocumentConverter;
import com.artofsolving.jodconverter.DocumentFormat;
import com.artofsolving.jodconverter.openoffice.connection.OpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.connection.SocketOpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.converter.StreamOpenOfficeDocumentConverter;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageTree;
import org.apache.pdfbox.rendering.ImageType;
import org.apache.pdfbox.rendering.PDFRenderer;
import org.apache.pdfbox.tools.imageio.ImageIOUtil;

import java.awt.image.BufferedImage;

import com.sun.net.httpserver.HttpExchange;
import com.sun.net.httpserver.HttpHandler;
import com.sun.net.httpserver.HttpServer;

import java.io.UnsupportedEncodingException;
import java.net.InetSocketAddress;
import java.net.URLDecoder;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;


public class OfficeHelper {

    public static void main(String[] args) throws Exception {
        //file2png("D:\\", "test.docx");
        HttpServer server = HttpServer.create(new InetSocketAddress(8001), 0);
        server.createContext("/office_convert", new OfficeConvertHandler());
        server.start();
        System.out.println("Server is listen on : 8001 ");
        // startOpenOfficeService();
        // file2png("D:\\", "test.xlsx");
        // shutdownOpenOfficeService();
    }

    static class OfficeConvertHandler implements HttpHandler{
        @Override
        public void handle(HttpExchange exchange) {
            new Thread(new Runnable() {
                @Override
                public void run() {
                    try{
                        //获得查询字符串(get)
                        String queryString =  exchange.getRequestURI().getQuery();
                        Map<String,String> queryStringInfo = formData2Dic(queryString);
                        //获得表单提交数据(post)
                        //String postString = IOUtils.toString(exchange.getRequestBody());
                        //Map<String,String> postInfo = formData2Dic(postString);
                        String fileDir = queryStringInfo.get("fileDir");
                        String fileName = queryStringInfo.get("fileName");

                        int response = file2png(fileDir, fileName);

                        exchange.sendResponseHeaders(200,0);
                        OutputStream os = exchange.getResponseBody();
                        os.write(String.valueOf(response).getBytes());
                        os.close();
                    }catch (IOException ie) {
                        ie.printStackTrace();
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                }
            }).start();
        }
    }

    public static Map<String,String> formData2Dic(String formData ) {
        Map<String,String> result = new HashMap<>();
        if(formData== null || formData.trim().length() == 0) {
            return result;
        }
        final String[] items = formData.split("&");
        Arrays.stream(items).forEach(item ->{
            final String[] keyAndVal = item.split("=");
            if( keyAndVal.length == 2) {
                try{
                    final String key = URLDecoder.decode( keyAndVal[0],"utf8");
                    final String val = URLDecoder.decode( keyAndVal[1],"utf8");
                    result.put(key,val);
                }catch (UnsupportedEncodingException e) {}
            }
        });
        return result;
    }

    public static int file2png(String fileDir, String fileName) throws Exception {
        if(fileDir == null || fileName == null){
            return 0;
        }

        // 输入文件
        File officeFile = new File(fileDir + File.separatorChar + fileName);
        // 临时文件输出路径
        String tempFileDir = fileDir + File.separatorChar + "temp" + File.separatorChar;
        // 判断是否为pdf
        String fileType = fileName.substring(fileName.lastIndexOf(".") + 1,fileName.length());

        boolean isTempPdf = false;
        File pdfFile = officeFile;
        if(!fileType.equals("pdf") && !fileType.equals("PDF")){
            // 转换为pdf
            isTempPdf = true;
            pdfFile = file2pdf(officeFile,tempFileDir);
            if(pdfFile == null){
                return 0;
            }
        }

        // 生成图片后的路径
        String pngFileDir = fileDir + File.separatorChar + fileName + "_png";
        // 生成图片
        return pdf2png(pdfFile,pngFileDir,isTempPdf);
    }

    public static File file2pdf(File inFile, String outDir)throws Exception{

        System.out.println("---------------office转pdf开始---" + outDir + File.separatorChar + inFile.getName());

        File dir = new File(outDir);
        if(!dir.exists()){
            dir.mkdirs();
        }

        String timesuffix = String.valueOf(System.currentTimeMillis());
        String tempOutFileName = timesuffix.concat(".pdf");

        File tempOutputFile = new File(outDir + File.separatorChar + tempOutFileName);
        if (tempOutputFile.exists()) {
            tempOutputFile.delete();
        }

        // 连接OpenOffice服务
        OpenOfficeConnection connection = new SocketOpenOfficeConnection("127.0.0.1", 8100);
        connection.connect();
        // convert
        DocumentConverter converter = new StreamOpenOfficeDocumentConverter(connection);
        DefaultDocumentFormatRegistry formatReg = new DefaultDocumentFormatRegistry();
        DocumentFormat odt = formatReg.getFormatByFileExtension("odt") ;
        DocumentFormat pdf = formatReg.getFormatByFileExtension("pdf") ;
        try {
            InputStream tempInputFileStream = new FileInputStream(inFile);
            OutputStream tempOutputFileStream = new FileOutputStream(tempOutputFile);
            converter.convert(tempInputFileStream,odt,tempOutputFileStream,pdf);
            tempInputFileStream.close();
            tempOutputFileStream.close();
        }catch (Exception e){
            e.printStackTrace();
            return null;
        }
        connection.disconnect();
        System.out.println("---------------office转pdf完成---" + outDir + File.separatorChar + inFile.getName());
        return tempOutputFile;
    }

    public static int pdf2png(File pdfFile, String outDir,boolean isTempPdf)throws Exception{
        System.out.println("---------------pdf转换png开始---" + outDir + File.separatorChar + pdfFile.getName());

        File dir = new File(outDir);
        if(!dir.exists()){
            dir.mkdirs();
        }

        // 生成图片计数
        int pageCounter = 0;
        // 生成图片
        try {
            PDDocument document = PDDocument.load(pdfFile);
            PDPageTree list = document.getDocumentCatalog().getPages();

            for (PDPage page : list) {
                PDFRenderer pdfRenderer = new PDFRenderer(document);
                BufferedImage image = pdfRenderer.renderImageWithDPI(pageCounter, 200, ImageType.RGB);
                String target = outDir + File.separatorChar + (pageCounter++) + ".png";
                ImageIOUtil.writeImage(image, target, 200);
                System.out.println(pdfFile.getName() + pageCounter);
            }
            document.close();
        }catch (Exception e){
            e.printStackTrace();
            return 0;
        }

        if(isTempPdf){
            //临时输出文件删除
            pdfFile.delete();
        }

        System.out.println("---------------pdf转换png完成---" + outDir + File.separatorChar + pdfFile.getName());
        return pageCounter;
    }

    /**
     * 启动服务
     **/
    public static void startOpenOfficeService() {
        String command = "C:\\Program Files (x86)\\OpenOffice 4\\program\\soffice -headless -accept=\"socket,host=127.0.0.1,port=8100;urp;\" -nofirststartwizard";
        try {
            Process pro = Runtime.getRuntime().exec(command);
        } catch (IOException e) {
            System.out.println("OpenOffice服务启动失败");
        }
    }

    /**
     * 关闭服务
     **/
    public static void shutdownOpenOfficeService() {
        Scanner in = null;
        try {
            Process pro = Runtime.getRuntime().exec("tasklist");
            in = new Scanner(pro.getInputStream());
            while (in.hasNext()) {
                String proString = in.nextLine();
                if (proString.contains("soffice.exe")) {
                    String cmd = "taskkill /f /im soffice.exe";
                    pro = Runtime.getRuntime().exec(cmd);
                    System.out.println("soffice.exe关闭");
                }
                if (proString.contains("soffice.bin")) {
                    String cmd = "taskkill /f /im soffice.bin";
                    pro = Runtime.getRuntime().exec(cmd);
                    System.out.println("soffice.bin关闭");
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (in != null) {
                in.close();
            }
        }
    }

}