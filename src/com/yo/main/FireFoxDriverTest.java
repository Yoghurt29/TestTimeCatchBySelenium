package com.yo.main;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.PrintWriter;
import java.net.URL;
import java.net.URLDecoder;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;


import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.TimeoutException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxBinary;
import org.openqa.selenium.firefox.FirefoxProfile;
import org.openqa.selenium.htmlunit.HtmlUnitDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import com.gargoylesoftware.htmlunit.BrowserVersion;
import com.yo.util.ExcelUtil;

import net.sf.json.JSONArray;
import net.sf.json.JSONObject;

/**
 * 基於WebDriver,抓取網頁 必須使用firefoxDriver 選擇器的使用,firefox唯一選擇器 注意使用延時
 * 如何將選擇器轉換成dom操作呢? vba的元素選擇方式太雞肋,且不能執行所有js,難以抓取部分動態加載頁面 WebDriver 2016.9.6
 * 简要测试，待繼續學習 2017.7.29编写最后一次测试时间抓取功能
 * 
 * @author Trulon_Chu
 */
public class FireFoxDriverTest {
	private static ArrayList<SnInfo> needCatchSnInfos = new ArrayList<SnInfo>();
	private static ArrayList<SnInfo> doneCatchSnInfos = new ArrayList<SnInfo>();
	// 使用这个构造可以呈现浏览器界面
	//private static File path=new File("D:\\Program Files (x86)\\Mozilla Firefox\\firefox.exe");
	//private static WebDriver driver =new org.openqa.selenium.firefox.FirefoxDriver(new FirefoxBinary(path),new FirefoxProfile());
	// 使用这个构造不呈现浏览器界面，注意使用这个构造需手动启用JS执行引擎（在静态代码块中）
	private static WebDriver driver = new HtmlUnitDriver(BrowserVersion.FIREFOX_24);
	private static JavascriptExecutor globalJavascriptExecutor = (JavascriptExecutor) driver;
	// 所有的js代码
	private static String jsCode = "";
	// 加载任意页面时候，等待该时间，还未找到页面上的某个元素则重试，默认为0，每次重试+1,最大MAX_WAIT_TIME_SECONDS
	private static int waitTimeSeconds = 0;
	private static int MAX_WAIT_TIME_SECONDS = 20;
	// 对列表中未抓取到记录的sn的进行全部重试的轮数计数
	private static int retryAllCount = 0;
	// 对列表中未抓取到记录的sn的进行全部重试的轮数的默认值
	private static int MAX_RETRY_ALL_COUNT = 2;
	static {
		((HtmlUnitDriver) driver).setJavascriptEnabled(true);
		try {
			jsCode = readJsCode();
		} catch (IOException e) {
			e.printStackTrace();
		}
		System.out.println("本次工作目录： "+getJarWorkPath());
		//每次清空已有文件
		File logFile =new File(getJarWorkPath()+"\\log.txt");
		File failExcel =new File(getJarWorkPath()+"抓取成功.xlsx");
		File doneExcel =new File(getJarWorkPath()+"抓取成功.xlsx");
		if(logFile.exists()){
			logFile.delete();
		}
		if(failExcel.exists()){
			failExcel.delete();
		}
		if(doneExcel.exists()){
			doneExcel.delete();
		}
	}

	public static void main(String[] args) throws Exception {
		initSns();
		startCatch();
		saveResult();
		uploadResult();
	}

	private static void uploadResult() throws InterruptedException {
		// com.yo.controller.InOutController public Map
		// uploadActiveTimeByGETRequestPlayloadJSONData(String
		// activeTimeJSONData,HttpServletRequest request)
		try {
			log("##準備上傳抓取的記錄至http://172.28.136.19:8099/PEMainboardTrack！請斷開VPN并稍等！正在重試。。。");
			JSONArray fromObject = JSONArray.fromObject(doneCatchSnInfos);
			
			//driver.get("http://172.28.136.19:8099/PEMainboardTrack/index.html");
			driver.get("http://172.28.136.19:8099/PEMainboardTrack/index.html");
			String pageSource = (String) globalJavascriptExecutor
					.executeScript(jsCode + "return uploadResultByAjaxPost('"+fromObject.toString()+"');");
			System.out.println(pageSource);
			//数据量大时候会失败
			/*driver.get(
					"http://172.28.136.19:8099/PEMainboardTrack/inOut/uploadActiveTimeByGETRequestPlayloadJSONData?activeTimeJSONData="
							+ fromObject.toString());*/
			System.out.println("pageSource: "+pageSource);
			if(pageSource.length()<100&&pageSource.contains("result")&&pageSource.contains("true")){
				log("似乎上传完成！请到http://172.28.136.19:8099/PEMainboardTrack/index.html核对本次上传结果!若失败，则请手动上传excel文件");
			}else{
				throw new RuntimeException("未能正確上傳！需重試！");
			}
		} catch (Exception e) {
			log("上傳失敗，請斷開VPN！準備重試！");
			for (int i = 0; i < 10; i++) {
				Thread.sleep(1000);
				log((10-i)+"秒后重試！");
			}
			uploadResult();
			return;
		}
	}

	private static void saveResult()
			throws NoSuchFieldException, SecurityException, IllegalArgumentException, IllegalAccessException {
		ExcelUtil.createExcelFile(doneCatchSnInfos, getJarWorkPath(), "抓取成功.xlsx",
				new String[] { "isn", "activeTime" }, new String[] { "isn", "activeTime" });
		ExcelUtil.createExcelFile(needCatchSnInfos, getJarWorkPath(), "抓取失败.xlsx",
				new String[] { "isn", "activeTime" }, new String[] { "isn", "activeTime" });
		log(
				doneCatchSnInfos.size() + "條SN測試記錄抓取成功！" + needCatchSnInfos.size() + "條SN測試記錄抓取失敗！\r\n 已保存至"+getJarWorkPath()+"目录下！");
	}

	public static void catchData()
			throws NoSuchFieldException, SecurityException, IllegalArgumentException, IllegalAccessException {
		Iterator<SnInfo> iterator = needCatchSnInfos.iterator();
		while (iterator.hasNext()) {
			SnInfo snInfo = iterator.next();
			try {
				globalJavascriptExecutor.executeScript(jsCode + "clickSerch('" + snInfo.getIsn() + "');");
				waitForElementExisit(driver,
						"form[name='f_13_1_5_15_1_3_0_1'] tbody:nth-child(1) tr:nth-child(2)  td:nth-child(17) font:nth-child(1)",
						waitTimeSeconds);
			} catch (TimeoutException e) {
				log(snInfo.getIsn() + "未查詢到該SN的測試記錄，跳過該SN，在下一輪再嘗試！");
				waitTimeSeconds=0;
				continue;
			}
			String timeStamp = (String) globalJavascriptExecutor
					.executeScript(jsCode + "return getLastTestTimestamp();");
			if (null != timeStamp) {
				snInfo.setActiveTime(timeStamp);
				log(snInfo+" 剩余 " + needCatchSnInfos.size());
				doneCatchSnInfos.add(snInfo);
				iterator.remove();
				// 应该有多次提交以保证中断后以前的不用再跑,但无法访问OA网络，下载时也下载了所有sn
				if(doneCatchSnInfos.size()%100==0){
					saveResult();
				}
			}
		}
		if (needCatchSnInfos.size() != 0 && retryAllCount < MAX_RETRY_ALL_COUNT) {
			retryAllCount++;
			waitTimeSeconds++;
			waitTimeSeconds = waitTimeSeconds > MAX_WAIT_TIME_SECONDS ? MAX_WAIT_TIME_SECONDS : waitTimeSeconds;
			log("有跳過的SN，還未完成！對未完成的SN進行下一輪！");
			log("##第" + (1 + retryAllCount) + "輪");
			throw new RuntimeException("有跳過的SN，還未完成！對未完成的SN進行下一輪！");
		} else {
			log("##抓取结束！ 总共嘗試輪數："+ (1 + retryAllCount) +(needCatchSnInfos.size()!=0?(" 但还有"+ needCatchSnInfos.size() + " 條sn 未抓取到測試記錄!!! " 
					+ " 可能是這些板子沒有測試記錄。"):""));
		}
	};

	public static void startCatch() {
		loginAndGoProductHistory();
		try {
			catchData();
		} catch (Exception e) {
			// e.printStackTrace();
			log("从QCR上抓取数据失败，请注意连接VPN！正在重试...");
			startCatch();
		}
	}

	public static void initSns() throws Exception {
		// 从excel获取
		 /*File fileExcel=new File("D:\\Users\\Trulon_Chu\\Desktop\\TestTimeCatch\\抓取成功.xlsx"); 
		 HashMap<Integer,String> config=new HashMap<>(); 
		 config.put(0, "isn"); 
		 List<SnInfo> list = ExcelUtil.readListFromExcel(fileExcel, SnInfo.class,config,new String[]{"isn","activeTime"});
		 doneCatchSnInfos.addAll(list);*/

		JSONObject fromObject = null;
		try {
			// 从主板管理系统获取所有SN 其中all 所有SN danger 超时SN notify 即将超时SN，详情参考如下接口
			// com.yo.controller.InOutController public Map
			// lendingBoardReport(@RequestParam(required=false)String
			// isn,@RequestParam(required=false)String
			// toPeopleWorkId,@RequestParam(required=false)String
			// takeTimeStart,@RequestParam(required=false)String
			// takeTimeEnd,@RequestParam(required=false)String isDelay,int
			// pageIndex,int capacity)
			log(
					"準備從 http://172.28.136.19:8099/PEMainboardTrack下載SN，若已連接VPN，斷開后等待幾秒，程序會繼續運行。請斷開VPN！正在等待網絡...");
			driver.get(
					"http://172.28.136.19:8099/PEMainboardTrack/inOut/lendingBoardReport?isDelay=all&pageIndex=1&capacity=-1");
			String pageSource = driver.getPageSource();
			fromObject = JSONObject.fromObject(pageSource);
		} catch (Exception e) {
			initSns();
			return;
		}
		JSONArray jsonArray = fromObject.getJSONObject("data").getJSONArray("list");
		List<SnInfo> list = new ArrayList<>();
		for (Object object : jsonArray) {
			JSONObject jsonObject = JSONObject.fromObject(object);
			list.add(new SnInfo(jsonObject.getString("isn"), null));
		}

		needCatchSnInfos.addAll(list);
		log("## " + needCatchSnInfos.size() + " 條sn下載完成， 準備從QCR上抓取數據，請連接VPN！正在等待網絡...");
	}

	public static void loginAndGoProductHistory() {
		try {
			driver.get("http://17.239.228.36/cgi-bin/WebObjects/QCR.woa/wa/default");
			waitForElementExisit(driver, "input[name='3.7.5.13']", waitTimeSeconds);
		} catch (TimeoutException e) {
			log("進入首頁失敗，正在重試...");
			loginAndGoProductHistory();
			return;
		}
		try {
			globalJavascriptExecutor.executeScript(jsCode + "login();");
			waitForElementExisit(driver, "#product_history", waitTimeSeconds);
		} catch (TimeoutException e) {
			log("登陸失敗，正在重試...");
			loginAndGoProductHistory();
			return;
		}
		try {
			globalJavascriptExecutor.executeScript(jsCode + "goProductHistory();");
			waitForElementExisit(driver, "input[name='7.1.3']", waitTimeSeconds);
		} catch (TimeoutException e) {
			log("進入ProductHistory失敗！正在重試...");
			loginAndGoProductHistory();
			return;
		}
	}

	public static String readJsCode() throws IOException {
		//File jsFile = new File(getJarWorkPath()+"D:\\workspace\\WebDriverDemo\\resource\\JS.js");
		File jsFile = new File(getJarWorkPath()+"\\JS.js");
		FileReader f = new FileReader(jsFile);
		BufferedReader br = new BufferedReader(f);
		String jarPath;
		String code = "";
		while (null != (jarPath = br.readLine())) {
			code += jarPath;
		}
		;
		return code;
	}

	public static void waitForElementExisit(WebDriver driver, String cssPickerExpress, int seconds) {
		WebDriverWait wait = new WebDriverWait(driver, seconds);
		wait.until(ExpectedConditions.presenceOfElementLocated(By.cssSelector(cssPickerExpress)));
	}

	public static String getJarWorkPath() {
		URL url = FireFoxDriverTest.class.getProtectionDomain().getCodeSource().getLocation();
		String filePath = null;
		try {
			filePath = URLDecoder.decode(url.getPath(), "utf-8");// 转化为utf-8编码
		} catch (Exception e) {
			e.printStackTrace();
		}
		if (filePath.endsWith(".jar")) {// 可执行jar包运行的结果里包含".jar"
			// 截取路径中的jar包名
			filePath = filePath.substring(0, filePath.lastIndexOf("/") + 1);
		}

		File file = new File(filePath);

		// /If this abstract pathname is already absolute, then the pathname
		// string is simply returned as if by the getPath method. If this
		// abstract pathname is the empty abstract pathname then the pathname
		// string of the current user directory, which is named by the system
		// property user.dir, is returned.
		filePath = file.getAbsolutePath();// 得到windows下的正确路径
		return filePath;
	}

	@SuppressWarnings("deprecation")
	public static void log(String log) {
		System.out.println(log);
		File f =null;
		PrintWriter pw=null;
		try {
			f = new File(getJarWorkPath()+"\\log.txt");
			pw = new PrintWriter(new BufferedWriter(new OutputStreamWriter(new FileOutputStream(f, true), "UTF-8")),false);
		} catch (IOException e) {
			e.printStackTrace();
		}
		pw.println((new Date()).toLocaleString()+"  "+log);
		pw.flush();
		pw.close();
	}
}
