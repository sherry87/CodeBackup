package test;

import java.io.IOException;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import javax.annotation.Resource;

import org.apache.commons.lang3.StringUtils;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.test.annotation.Rollback;
import org.springframework.test.context.ContextConfiguration;
import org.springframework.test.context.junit4.SpringJUnit4ClassRunner;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.execl.ImportExecl;
import com.hintsoft.agent.test.autoGenerateTool;

@RunWith(SpringJUnit4ClassRunner.class)
@ContextConfiguration("classpath:applicationContext.xml")
@Rollback(false)
public class ImportExcelTest {
	@Resource
	JdbcTemplate jdbcTemplate;
	public String url="https://sherry87.github.io/resume/";

	
	@Test
	public void test() throws Exception{
		/*ImportExecl poi = new ImportExecl();
		 List<List<List<String>>> list = poi.read("E:\\1.xlsx");
		 List<List<List<String>>> list2 = poi.read("E:\\2.xls");
		 list.addAll(list2);
		 if (list != null) {
	            for (int i = 0; i < list.size(); i++) {
	            	saveSheet(list.get(i));
	            }
	    }*/
		 
		// generateHTML();
		 generateEwmJson();
		generateHTMLInfo();
	}
	
	
	
	/**
	 * 生成HTML
	 * @throws IOException 
	 */
	public void generateHTML() throws IOException{
		 ImportExecl poi = new ImportExecl();
		 List<List<List<String>>> list = poi.read("E:\\1.xlsx");
		 List<List<List<String>>> list2 = poi.read("E:\\2.xls");
		 list.addAll(list2);
		 if (list != null) {
	            for (int i = 0; i < list.size(); i++) {
	            	Map<String,String> map= sheetToMap(list.get(i));
	            	generateFile(map);
	            }
	    }
	}
	
	/**
	 * 生成手机端HTML
	 * @throws IOException 
	 */
	public void generateHTMLInfo() throws IOException{
		 ImportExecl poi = new ImportExecl();
		 List<List<List<String>>> list = poi.read("E:\\1.xlsx");
		 List<List<List<String>>> list2 = poi.read("E:\\2.xls");
		 list.addAll(list2);
		 if (list != null) {
	            for (int i = 0; i < list.size(); i++) {
	            	Map<String,String> map= sheetToMap(list.get(i));
	            	generateFileInfo(map);
	            }
	    }
	}
	
	/**
	 * 封装sheet为map
	 * @param dataList
	 * @return
	 */
	public Map<String,String> sheetToMap(List<List<String>> dataList){
		Map<String,String> map = new LinkedHashMap<String,String>();
	   	 for (int j = 0; j < dataList.size(); j++) {
	       	 List<String> cellList = dataList.get(j);
	       	 map.put(cellList.get(0).replaceAll("：", ""), cellList.get(1));
	       }
	   	 return map;
	}
	
	public void generateFile(Map<String,String> map) throws IOException{
		StringBuffer str1 = new StringBuffer("");
		boolean second=false;
		StringBuffer str2 = new StringBuffer("");
		boolean third = false;
		StringBuffer str3 = new StringBuffer("");
		
		for(Entry<String,String> entry:map.entrySet()){
			if(entry.getKey().indexOf("在线验证码")>-1){
				second=false;
				third=true;
			}
			if(second==false&&third==false){
				str1.append("<div class=\'row collection-item\'><div class=\'col s4 bolder\'>"+entry.getKey()+"：</div><div class=\'col s8\'>"+entry.getValue()+"</div></div>").append("\n");
				if(entry.getKey().indexOf("出生日期")>-1){
					second=true;
				}
			}else if(second){
				str2.append("<div class=\'row collection-item\'><div class=\'col s4 bolder\'>"+entry.getKey()+"：</div><div class=\'col s8\'>"+entry.getValue()+"</div></div>").append("\n");
			}else if(third){
				str3.append("<div class=\'row collection-item\'><div class=\'col s4 bolder\'>"+entry.getKey()+"：</div><div class=\'col s8\'>"+entry.getValue()+"</div></div>").append("\n");
			}
		}
		
		String yzm = map.get("在线验证码").replaceAll(" ", "");
		StringBuffer sb = new StringBuffer("");
		sb.append("<!doctype html>").append("\n");
		sb.append("<html>").append("\n");
		sb.append("<head>").append("\n");
		sb.append("<meta charset=\'utf-8\'/>").append("\n");
		sb.append("<meta name=\'viewport\' content=\'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=0\' />").append("\n");
		sb.append("<meta name=\'format-detection\' content=\'telephone=no\' />").append("\n");
		sb.append("<title>教育部学历证书电子注册备案表</title>").append("\n");
		sb.append("<link href=\'https://cdn.bootcss.com/materialize/0.99.0/css/materialize.css\' rel=\'stylesheet\'>").append("\n");
		sb.append("<link rel=\'stylesheet\' href=\'css/wap_md.min.css\'/>").append("\n");
		sb.append("</head>").append("\n");
		sb.append("").append("\n");
		sb.append("<body>").append("\n");
		sb.append("<div class=\'page-wrapper\'>").append("\n");
		sb.append("  <div class=\'pw-content\'>").append("\n");
		sb.append("    <div class=\'container\'>").append("\n");
		sb.append("        <div class=\'row valign-wrapper padTop\' id=\'xl-top\'>").append("\n");
		sb.append("					<div class=\'col s3 \'>").append("\n");
		sb.append("					    <img src=\'img/search_pdf.png\' alt=\'下载PDF\' title=\'下载PDF\' class=\'responsive-img waves-effect waves-teal right z-depth-1 down_psf\'>").append("\n");
		sb.append("					</div>").append("\n");
		sb.append("        </div>").append("\n");
		sb.append("        <h5 class=\'center-align padBot\'>教育部学历证书电子注册备案表</h5>").append("\n");
		sb.append("        <div class=\'row\'>").append("\n");
		sb.append("            <div class=\'col s12 center-align\' style=\'min-height:11rem;\'>").append("\n");
		sb.append("                <img class=\'waves-effect waves-teal z-depth-1 h_img_size\' id=\'xueliPhoto\' src=\'\' onerror=\'imgErr()\' data-caption=\'个人头像\' />").append("\n");
		sb.append("            </div>").append("\n");
		sb.append("        </div>").append("\n");
		sb.append("    </div>").append("\n");
		sb.append("").append("\n");
		sb.append("    <div class=\'collection card\'>").append("\n");
		sb.append(str1.toString());
		sb.append("    </div>").append("\n");
		
		sb.append("    <div class=\'collection card\'>").append("\n");
		sb.append(str2.toString());
		sb.append("    </div>").append("\n");
		
		sb.append("    <div class=\'collection card\'>").append("\n");
		sb.append(str3.toString());
		sb.append("    </div>").append("\n");
		
		sb.append("").append("\n");
		sb.append("").append("\n");
		sb.append("    <div class=\'container\'>").append("\n");
		sb.append("        <div class=\'row\'>").append("\n");
		sb.append("            <div class=\'col s12\'> ").append("\n");
		sb.append("            </div>").append("\n");
		sb.append("        </div>").append("\n");
		sb.append("    </div>").append("\n");
		sb.append("  </div>").append("\n");
		sb.append("</div>").append("\n");
		sb.append("</body>").append("\n");
		sb.append("</html>").append("\n");
		sb.append("<script src=\'https://cdn.bootcss.com/jquery/1.12.1/jquery.min.js\'></script>").append("\n");
		sb.append("<!--<script type=\'text/javascript\' src=\'js/materialize.min.js\'></script> -->").append("\n");
		sb.append("<script src=\'https://cdn.bootcss.com/materialize/0.99.0/js/materialize.min.js\'></script>").append("\n");
		sb.append("").append("\n");
		
		sb.append("<script>").append("\n");
		sb.append("$(function(){").append("\n");
		sb.append("		$(\'#xueliPhoto\').attr(\'src\',\'img/"+yzm+".jpg\');").append("\n");
		sb.append("});    ").append("\n");
		sb.append("	function imgErr(){");
		sb.append("		$(\'#xueliPhoto\').hide();");
		sb.append("	}");
		sb.append("</script>").append("\n");
		
		
		autoGenerateTool.createFile("E:\\", yzm+".html", sb.toString());
	}
	
	//生成web端详情页面
	public void generateFileInfo(Map<String,String> map) throws IOException{
		String yzm = map.get("在线验证码").replaceAll(" ", "");
		StringBuffer sb = new StringBuffer("");
		sb.append("<!DOCTYPE html> \n ");
		sb.append("<html> \n ");
		sb.append("	<head> \n ");
		sb.append("		<meta http-equiv=\'Content-Type\' content=\'text/html; charset=UTF-8\' /> \n ");
		sb.append("		<title>教育部学历证书电子注册备案表</title> \n ");
		sb.append("		<link href=\'../css/common.css\' rel=\'stylesheet\' type=\'text/css\' /> \n ");
		sb.append("	</head> \n ");
		sb.append("	<body> \n ");
		sb.append("		<div class=\'main clearfix\'> \n ");
		sb.append("			<div class=\'m_s_r\' id=\'rightCnt\'> \n ");
		sb.append("				<div class=\'m_cnt_l\'> \n ");
		sb.append("					<div id=\'resultTable\'> \n ");
		sb.append("						<div class=\'tableTitle\'>教育部学历证书电子注册备案表</div> \n ");
		sb.append("						<div class=\'div1\'> \n ");
		sb.append("							<div class=\'div2\'> \n ");
		sb.append("								<table width=\'628\' border=\'0\' align=\'center\' cellpadding=\'0\' cellspacing=\'0\' class=\'cn_table\'> \n ");
		sb.append("									<col width=\'91\'> \n ");
		sb.append("									<col width=\'172\'> \n ");
		sb.append("									<col width=\'91\'> \n ");
		sb.append("									<col width=\'140\'> \n ");
		sb.append("									<col width=\'133\'> \n ");
		sb.append("									<tr> \n ");
		sb.append("										<th>姓　　名</th> \n ");
		sb.append("										<td colspan=\'3\'>"+map.get("姓名")+"</td> \n ");
		sb.append("										<td rowspan=\'4\'> \n ");
		sb.append("											<div class=\'cn_photo1\' id=\'xueli_photo_div\'><img class=\'cn_photo1_img\' id=\'xueliPhoto\' src=\'../img/"+yzm+".jpg\' width=\'120.0\' height=\'160.0\' /></div> \n ");
		sb.append("										</td> \n ");
		sb.append("									</tr> \n ");
		sb.append("									<tr> \n ");
		sb.append("										<th> \n ");
		sb.append("											<div style=\'width:90px;\'>性　　别</div> \n ");
		sb.append("										</th> \n ");
		sb.append("										<td><span id=\'xb\' style=\'width:166px;\'>"+map.get("性别")+"</span></td> \n ");
		sb.append("										<th> \n ");
		sb.append("											<div style=\'width:90px;\'>出生日期</div> \n ");
		sb.append("										</th> \n ");
		sb.append("										<td><span style=\'width:134px;\'>"+map.get("出生日期")+"</span></td> \n ");
		sb.append("									</tr> \n ");
		sb.append("									<tr> \n ");
		sb.append("										<th>入学时间</th> \n ");
		sb.append("										<td><span>"+formate(map.get("入学时间"))+"</span></td> \n ");
		sb.append("										<th>毕业时间</th> \n ");
		sb.append("										<td><span>"+formate(map.get("毕业时间"))+"</span></td> \n ");
		sb.append("									</tr> \n ");
		sb.append("									<tr> \n ");
		sb.append("										<th>学历类型</th> \n ");
		sb.append("										<td><span>"+map.get("学历类型")+"</span></td> \n ");
		sb.append("										<th>学历层次</th> \n ");
		sb.append("										<td><span>"+map.get("学历层次")+"</span></td> \n ");
		sb.append("									</tr> \n ");
		sb.append("									<tr> \n ");
		sb.append("								</table> \n ");
		sb.append("								<table width=\'628\' border=\'0\' align=\'center\' cellpadding=\'0\' cellspacing=\'0\' class=\'cn_table cn_table_noborder\'> \n ");
		sb.append("									<col width=\'91\'> \n ");
		sb.append("									<col width=\'302\'> \n ");
		sb.append("									<col width=\'101\'> \n ");
		sb.append("									<col width=\'133\'> \n ");
		sb.append("									<tr> \n ");
		sb.append("										<th> \n ");
		sb.append("											<div style=\'width:90px;\'>毕业院校</div> \n ");
		sb.append("										</th> \n ");
		sb.append("										<td><span style=\'width:296px;\'>"+map.get("毕业院校")+"</span></td> \n ");
		sb.append("										<th> \n ");
		sb.append("											<div style=\'width:100px;\'>院校所在地</div> \n ");
		sb.append("										</th> \n ");
		sb.append("										<td><span style=\'width:128\'>"+formate(map.get("院校所在地"))+"</span></td> \n ");
		sb.append("									</tr> \n ");
		sb.append("									<tr> \n ");
		sb.append("										<th>专业名称</th> \n ");
		sb.append("										<td><span>"+map.get("专业名称")+"</span></td> \n ");
		sb.append("										<th class=\'cn_font1\'>学习形式</th> \n ");
		sb.append("										<td><span>"+formate(map.get("学习形式"))+"</span></td> \n ");
		sb.append("									</tr> \n ");
		sb.append("									<tr> \n ");
		sb.append("										<th>证书编号</th> \n ");
		sb.append("										<td class=\'cn_font2\'><span>"+map.get("证书编号")+"</span></td> \n ");
		sb.append("										<th>毕结业结论</th> \n ");
		sb.append("										<td><span>"+map.get("毕结业结论")+"</span></td> \n ");
		sb.append("									</tr> \n ");
		sb.append("									<tr> \n ");
		sb.append("										<th rowspan=\'3\'>二<br />维<br />验<br />证<br />码</th> \n ");
		sb.append("										<td rowspan=\'3\'> \n ");
		sb.append("											<div class=\'cn_photo2\' id=\'ewm\'> \n ");
		sb.append("												<img class=\'cn_photo2_txm\' src=\'../img/ewm2.jpg\' width=\'159\' height=\'123\' /> \n ");
		sb.append("											</div> \n ");
		sb.append("										</td> \n ");
		sb.append("										<th>在线验证码</th> \n ");
		sb.append("										<td class=\'cn_font2\'><span>"+map.get("在线验证码")+"</span></td> \n ");
		sb.append("									</tr> \n ");
		sb.append("									<tr> \n ");
		sb.append("										<th class=\'cn_font1\'>制表日期</th> \n ");
		sb.append("										<td><span>"+formate(map.get("制表日期"))+"</span></td> \n ");
		sb.append("									</tr> \n ");
		sb.append("									<tr> \n ");
		sb.append("										<th class=\'cn_font1\'>验证期至</th> \n ");
		sb.append("										<td><span>"+formate(map.get("验证期至"))+"</span></td> \n ");
		sb.append("									</tr> \n ");
		sb.append("									<tr> \n ");
		sb.append("										<td colspan=\'4\'> \n ");
		sb.append("											<h2>注意事项：</h2> \n ");
		sb.append("											<div class=\'zysx\'> \n ");
		sb.append("												<table width=\'100%\' border=\'0\' cellspacing=\'0\' cellpadding=\'0\'> \n ");
		sb.append("													<tr> \n ");
		sb.append("														<td valign=\'top\'>1、</td> \n ");
		sb.append("														<td valign=\'top\'>备案表是依据《高等学校学生学籍学历电子注册办法》（ \n ");
		sb.append("															<a href=\'http://www.chsi.com.cn/jyzx/201408/20140829/1245955796.html\' target=\'_blank\' style=\'text-decoration:underline;\'>教学[2014]11号</a>）对学历证书电子注册复核备案的结果；由教育部指定的唯一学历查询网站中国高等教育学生信息网（ \n ");
		sb.append("															<a href=\'http://www.chsi.com.cn\' target=\'_blank\'>http://www.chsi.com.cn</a>）提供在线验证服务。</td> \n ");
		sb.append("													</tr> \n ");
		sb.append("													<tr> \n ");
		sb.append("														<td valign=\'top\'>2、</td> \n ");
		sb.append("														<td valign=\'top\'>备案表内容验证办法：①点击备案表(电子版)中的在线验证码，可在线验证；②登录中国高等教育学生信息网“在线验证系统”，输入在线验证码进行验证；③利用专业扫描工具或具有条码识别功能的手机，扫描备案表中的二维码进行验证。</td> \n ");
		sb.append("													</tr> \n ");
		sb.append("													<tr> \n ");
		sb.append("														<td valign=\'top\'>3、</td> \n ");
		sb.append("														<td valign=\'top\'>备案表在验证有效期内可免费打印和验证。</td> \n ");
		sb.append("													</tr> \n ");
		sb.append("													<tr> \n ");
		sb.append("														<td valign=\'top\'>4、</td> \n ");
		sb.append("														<td valign=\'top\'>备案表内容如有修改，请以最新在线验证的内容为准。</td> \n ");
		sb.append("													</tr> \n ");
		sb.append("													<tr> \n ");
		sb.append("														<td valign=\'top\'>5、</td> \n ");
		sb.append("														<td valign=\'top\'>备案表内容标注“＊”号，表示学历信息该项内容不详。</td> \n ");
		sb.append("													</tr> \n ");
		sb.append("													<tr> \n ");
		sb.append("														<td valign=\'top\'>6、</td> \n ");
		sb.append("														<td valign=\'top\'>未经学历信息权属人同意，不得将备案表用于违背权属人意愿之用途。</td> \n ");
		sb.append("													</tr> \n ");
		sb.append("												</table> \n ");
		sb.append("											</div> \n ");
		sb.append("										</td> \n ");
		sb.append("									</tr> \n ");
		sb.append("								</table> \n ");
		sb.append("							</div> \n ");
		sb.append("						</div> \n ");
		sb.append(" \n ");
		sb.append("					</div> \n ");
		sb.append("				</div> \n ");
		sb.append("			</div> \n ");
		sb.append("		</div> \n ");
		sb.append("		<script type=\'text/javascript\' src=\'../js/jquery-1.11.1.js\' ></script> \n ");
		sb.append("		<script type=\'text/javascript\' src=\'../js/jquery.qrcode.js\' ></script> \n ");
		sb.append("        <script type=\'text/javascript\' src=\'../js/qrcode.js\' ></script>  \n ");
		sb.append("        <script type=\'text/javascript\' src=\'../js/utf.js\' ></script> \n ");
		sb.append("	</body> \n ");
		sb.append("</html> \n ");
		sb.append("<script type=\'text/javascript\'> \n ");
		sb.append("	$(function() { \n ");
		sb.append("		var url = \'"+url+yzm+".html\'; \n ");
		sb.append("					$(\'#ewm\').qrcode({ \n ");
		sb.append("				         text    : url ,  \n ");
		sb.append("				         width : \'123\', \n ");
		sb.append("		                 height : \'123\', \n ");
		sb.append("				     }); \n ");
		sb.append(" \n ");
		sb.append("		var imgLeft = ($(\'#xueli_photo_div\').width() - $(\'#xueliPhoto\').attr(\'width\')) / 2; \n ");
		sb.append("		var imgTop = ($(\'#xueli_photo_div\').height() - $(\'#xueliPhoto\').attr(\'height\')) / 2; \n ");
		sb.append("		$(\'#xueliPhoto\').css({ \n ");
		sb.append("			\'left\': imgLeft + \'px\', \n ");
		sb.append("			\'top\': imgTop + \'px\' \n ");
		sb.append("		}); \n ");
		sb.append("	}); \n ");
		sb.append("</script> \n ");
		autoGenerateTool.createFile("E:\\resume\\", "view-"+yzm+".html", sb.toString());
	}
	
	public  String formate(String txt){
		String r="";
		if(StringUtils.isNotEmpty(txt)){
				r=txt;
		}
		return r;
	}
	public void generateEwmJson(){
		 JSONArray arr = new JSONArray();
		 ImportExecl poi = new ImportExecl();
		 List<List<List<String>>> list = poi.read("E:\\1.xlsx");
		 List<List<List<String>>> list2 = poi.read("E:\\2.xls");
		 list.addAll(list2);
		 if (list != null) {
	            for (int i = 0; i < list.size(); i++) {
	            	Map<String,String> map= sheetToMap(list.get(i));
	            	String yzm = map.get("在线验证码").replaceAll(" ", "");
	            	JSONObject obj = new JSONObject();
	            	obj.put("name", map.get("姓名"));
	            	obj.put("sex", map.get("性别"));
	            	obj.put("code", yzm);
	            	obj.put("url", url+yzm+".html");
	            	arr.add(obj);
	            }
	    }
		 
		 try {
			autoGenerateTool.createFile("E:\\", "resume.json", arr.toJSONString());
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	/**
	 * 保存入库
	 * @param dataList
	 */
	public void saveSheet(List<List<String>> dataList){
		Map<String,String> map = new HashMap<String,String>();
        for (int j = 0; j < dataList.size(); j++) {
        	 List<String> cellList = dataList.get(j);
        	 int k=0;
    		 if(cellList.get(k).indexOf("姓名")>-1){
    			 map.put("USER_NAME", cellList.get(k+1));
    		 }else if(cellList.get(k).indexOf("性别")>-1){
    			 map.put("SEX", cellList.get(k+1));
    		 }else if(cellList.get(k).indexOf("出生日期")>-1){
    			 map.put("BIRTHDAY", cellList.get(k+1));
    		 }else if(cellList.get(k).indexOf("入学")>-1){
    			 map.put("JOIN_SCHOOL", cellList.get(k+1));
    		 }else if(cellList.get(k).indexOf("毕业时间")>-1){
    			 map.put("LEAVE_SCHOOL", cellList.get(k+1));
    		 }else if(cellList.get(k).indexOf("学历类型")>-1){
    			 map.put("XL_TYPE", cellList.get(k+1));
    		 }else if(cellList.get(k).indexOf("学历层次")>-1){
    			 map.put("XL_CC", cellList.get(k+1));
    		 }else if(cellList.get(k).indexOf("毕业院校")>-1){
    			 map.put("BYYX", cellList.get(k+1));
    		 }else if(cellList.get(k).indexOf("院校所在地")>-1){
    			 map.put("YXSZD", cellList.get(k+1));
    		 }else if(cellList.get(k).indexOf("专业名称")>-1){
    			 map.put("ZYMC", cellList.get(k+1));
    		 }else if(cellList.get(k).indexOf("学习形式")>-1){
    			 map.put("XXXS", cellList.get(k+1));
    		 }else if(cellList.get(k).indexOf("证书编号")>-1){
    			 map.put("ZSBH", cellList.get(k+1));
    		 }else if(cellList.get(k).indexOf("结业结论")>-1){
    			 map.put("BYJL", cellList.get(k+1));
    		 }else if(cellList.get(k).indexOf("验证码")>-1){
    			 map.put("ZXYZM", cellList.get(k+1));
    		 }else if(cellList.get(k).indexOf("制表日期")>-1){
    			 map.put("ZBRQ", cellList.get(k+1));
    		 }else if(cellList.get(k).indexOf("验证期至")>-1){
    			 map.put("YZRQ", cellList.get(k+1));
    		 }
        	
        }
        
	     String sql =generateSql(map);
	     jdbcTemplate.execute(sql);
	}
	
	public String generateSql(Map<String,String> map){
		String sql = " INSERT INTO `resume` ";
		StringBuffer sb = new StringBuffer("`ID`");
		StringBuffer sb2 = new StringBuffer("NULL");
		for (Map.Entry<String,String> entry : map.entrySet()) {
			sb.append(",`"+entry.getKey()+"`");
			sb2.append(",'"+entry.getValue()+"'");
	   	 }
		sql = sql +"("+sb+") VALUES ("+sb2+")";
		return sql;
	}
	
	public static void main(String[] args){
		String msg = "2017年07月27日";
		String m = msg.substring(5,7);
		String d = msg.substring(8,10);
		System.out.println(m);
		System.out.println(d);
		System.out.println(Integer.valueOf(m));
	}
	
}
