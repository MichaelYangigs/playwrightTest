package com.IGS;

import com.microsoft.playwright.*;
import com.microsoft.playwright.options.AriaRole;
import com.microsoft.playwright.options.LoadState;
import com.microsoft.playwright.options.WaitForSelectorState;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.BufferedReader;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Pattern;

public class Main {
    // CSV 列索引常量  根据test_data.csv 调整
    private static final int PRINCIPAL_NAME = 0;
    private static final int B_R_CERTIFICATE_NO = 1;
    private static final int ADDRESS = 2;
    private static final int CONTACT_NAME = 3;
    private static final int POST_TITLE = 4;
    private static final int TEL_NO = 5;
    private static final int EMAIL = 6;
    private static final int CONTRACT_NO = 7;
    private static final int CONTRACT_NAME = 8;
    private static final int WORK_SITE = 9;
    private static final int COMMENCEMENT_DATE = 10;
    private static final int COMPLETION_DATE = 11;
    private static final int IMPLEMENTATION_MEASURES = 12;
    private static final int TRADE = 13;
    private static final int PERIOD_START = 14;
    private static final int PERIOD_END = 15;
    private static final int MANPOWER_REQUIRED = 16;
    private static final int LOCAL_LABOUR = 17;
    private static final int IMPORTED_LABOUR = 18;
    private static final int TOTAL_MONTHS = 19;
    private static final int MONTHLY_WAGE = 20;
    private static final int SKILLED_WORKERS = 21;
    private static final int OTHERS = 22;
    private static final int IMPORTED_APPLIED = 23;
    private static final int APPROVED_No = 24;
    private static final int EMAIL_SUBJECT = 25;

    private static JFrame frame;
    private static JTextField testLinkField;
    private static JTextField csvPathField;
    private static JTextArea logArea;
    private static JButton startButton;

    public static void main(String[] args) {
        // 创建并显示GUI
        SwingUtilities.invokeLater(() -> createAndShowGUI());
    }

    private static void createAndShowGUI() {
        // 创建主窗口
        frame = new JFrame("Automated testing configuration");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(600, 400);
        frame.setLayout(new BorderLayout());

        // 创建输入面板
        JPanel inputPanel = new JPanel(new GridLayout(3, 2, 5, 5));
        inputPanel.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));

        // 添加测试链接输入
        inputPanel.add(new JLabel("Test link:"));
        testLinkField = new JTextField("https://localhost:2222/#/login"); // 默认
        inputPanel.add(testLinkField);

        // 添加CSV路径输入
        inputPanel.add(new JLabel("CSV file path:"));
        csvPathField = new JTextField("D://project/test_data.csv"); // 默认路径
        inputPanel.add(csvPathField);

        // 开始按钮
        startButton = new JButton("Start Testing");
        inputPanel.add(startButton);

        // 添加日志区域
        logArea = new JTextArea();
        logArea.setEditable(false);
        JScrollPane scrollPane = new JScrollPane(logArea);

        // 将组件添加到主窗口
        frame.add(inputPanel, BorderLayout.NORTH);
        frame.add(scrollPane, BorderLayout.CENTER);

        // 添加按钮事件监听器
        startButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                startButton.setEnabled(false);

                // 添加进度条
                JProgressBar progressBar = new JProgressBar(0, 100);
                progressBar.setStringPainted(true);
                frame.add(progressBar, BorderLayout.SOUTH);
                frame.revalidate();

                new Thread(() -> {
                    String testLink = testLinkField.getText();
                    String csvPath = csvPathField.getText();

                    // 立即更新UI表示开始
                    updateProgress(progressBar, 10, "Initializing browser...");

                    try {
                        runAutomation(testLink, csvPath, progressBar);
                    } catch (Exception ex) {
                        appendLog("An error occurred during automation: " + ex.getMessage());
                        ex.printStackTrace();
                    } finally {
                        SwingUtilities.invokeLater(() -> {
                            frame.remove(progressBar);
                            frame.revalidate();
                            startButton.setEnabled(true);
                        });
                    }
                }).start();
            }
        });

        // 显示窗口
        frame.setVisible(true);
    }

    private static void runAutomation(String testLink, String csvPath, JProgressBar progressBar) throws Exception {
        updateProgress(progressBar, 20, "Reading CSV data...");

        List<String[]> csvData = readCSV(csvPath);
        if (csvData.isEmpty()) {
            appendLog("CSV data not found, Exit the program.");
            return;
        }

        // 计算每个应用的进度增量
        final int BASE_PROGRESS = 60; // 基础进度（登录等操作完成后的进度）
        final int TOTAL_APPS_PROGRESS = 40; // 所有应用共享的进度区间
        double progressPerApp = (double) TOTAL_APPS_PROGRESS / csvData.size();

        updateProgress(progressBar, 30, "Launching browser...");

        try (Playwright playwright = Playwright.create()) {
            Browser browser = playwright.chromium().launch(
                    new BrowserType.LaunchOptions()
                            .setHeadless(false)
                            .setTimeout(120000)
            );

            updateProgress(progressBar, 40, "Creating browser context...");

            BrowserContext context = browser.newContext(
                    new Browser.NewContextOptions().setIgnoreHTTPSErrors(true)
            );

            updateProgress(progressBar, 50, "Opening new page...");

            Page page = context.newPage();

            //进度条更新
            updateProgress(progressBar, 60, "Logging in...");
            // 登录流程
            updateProgress(progressBar, 50, "Logging in...");
            login(page, testLink);

            int total = csvData.size();
            double currentProgress = BASE_PROGRESS;

            for (int i = 0; i < total; i++) {
                String[] row = csvData.get(i);

                // 更新进度：开始处理新应用
                updateProgress(progressBar, (int) currentProgress,
                        "Creating application " + (i + 1) + "/" + total);

                try {
                    // 创建Application
                    createApplication(page, row);
                    appendLog("Created Application: " + row[CONTRACT_NAME]);

                    // 更新进度：应用创建完成
                    currentProgress += progressPerApp * 0.3;
                    updateProgress(progressBar, (int) currentProgress,
                            "Processing workflow for application " + (i + 1) + "/" + total);

                    // 审批流程操作
                    processApplicationWorkflow(page, row);

                    // 更新进度：流程处理完成
                    currentProgress += progressPerApp * 0.7;
                    updateProgress(progressBar, (int) currentProgress,
                            "Completed application " + (i + 1) + "/" + total);

                    appendLog(">>>>>>>>>> The " + (i+1) + " data testing operation is completed ~ <<<<<<<<<<");

                } catch (Exception ex) {
                    appendLog("Error processing application " + (i + 1) + ": " + ex.getMessage());
                    throw ex;
                }
            }

            // 确保最终进度为100%
            updateProgress(progressBar, 100, "Automation completed!");
            appendLog("============== Automated testing completed! Created " + total + " applications. ==============");
        }
    }

    // 执行Application的审批工作流程
    private static void processApplicationWorkflow(Page page, String[] data) throws Exception {
        try {
            // 1. 提交申请
            submitApplication(page);

            // 导航回应用列表
            navigateToApplicationList(page);
            // 修改Approved No
            updateApprovedNo(page, data);

            // 2. 接受申请
            acceptApplication(page);

            // 3. 发送确认邮件
            issueAcknowledgement(page, data);

            // 4. 完成会议
            finishMeeting(page);

            // 5. 决策
            makeDecision(page);

            // 6. 请求批准
            requestApproval(page);

            // 7. 确认输入
            confirmInput(page);

            // 8. 发送邮件并生成quota
            draftNotification(page, data);

            String contractName = getFirstContractName(page);
            appendLog("Completed workflow for application: " + contractName);
        } catch (Exception e) {
            appendLog("Error processing workflow: " + e.getMessage());
            throw e;
        }
    }

    // 提交申请
    private static void submitApplication(Page page) throws Exception {
        // 点击请求批准按钮
        page.locator("button:has-text('Submit')").click();

        // 检查弹窗是否出现（2秒内）
        boolean isDialogPresent = false;
        try {
            page.waitForSelector(".el-message-box:visible",
                    new Page.WaitForSelectorOptions().setTimeout(2000));
            isDialogPresent = true;
        } catch (Exception e) {
            appendLog("No confirmation dialog appeared");
        }

        // 只有出现弹窗时才处理
        if (isDialogPresent) {
            handleConfirmationDialog(page);
            appendLog("Handled confirmation dialog");
        }

        // 等待操作完成
        page.waitForSelector(".el-message--success",
                new Page.WaitForSelectorOptions().setTimeout(5000));

        appendLog("Application submitted successfully");
    }

    private static void updateApprovedNo(Page page, String[] data) throws Exception {
        // 显式刷新列表
        navigateToApplicationList(page);

        // 获取Application list页面第一条数据点击编辑按钮
        getFirstTableRowAndClickEditButton(page);

        // 点击 Manpower Plan 标签页
        appendLog("Click Manpower Plan Tab...");
        page.getByRole(AriaRole.TAB, new Page.GetByRoleOptions().setName("Manpower Plan")).click();
        page.waitForTimeout(1500); // 等待内容加载

        // 定位Common Trade表格的第一行
        Locator commonTradeTable = page.locator(
                "//div[contains(@class, 'form-part-title') and contains(., 'Part 2: Manpower Requirement By Trade')]" +
                        "//following-sibling::div[contains(@class, 'form-item-content')]" +
                        "//div[contains(@class, 'table-title') and contains(., 'Part 2a: Skilled Worker')]" +
                        "//following-sibling::div[contains(@class, 'el-table') and .//th[contains(., 'Common Trade')]]"
        );

        // 获取第一行数据
        Locator firstRow = commonTradeTable.locator("//tbody/tr[1]");

        // 点击第一行的编辑按钮
        firstRow.locator("//td[contains(@class, 'operation')]//i[contains(@class, 'mdi-square-edit-outline')]").click();

        // 等待编辑弹窗出现
        page.waitForSelector(".el-dialog:visible", new Page.WaitForSelectorOptions().setTimeout(30000));

        // 填写Approved No.字段
        fillField(page, "Approved No. of quota:", data[APPROVED_No]);

        // 人力计划保存
        page.getByLabel("Manpower Requirement By Trade").getByRole(AriaRole.BUTTON, new Locator.GetByRoleOptions().setName("Save")).click();

        // 最后保存
        page.locator("#app").getByRole(AriaRole.BUTTON, new Locator.GetByRoleOptions().setName("Save")).click();
        page.waitForTimeout(1000); // 等待输入完成
    }

    // 导航到应用列表
    private static void navigateToApplicationList(Page page) throws Exception {
        page.locator(".el-menu-item:has-text('Applications')").click();
        page.waitForLoadState(LoadState.NETWORKIDLE);
        page.waitForSelector("button:has-text('Add Application')",
                new Page.WaitForSelectorOptions().setTimeout(30000));
        page.waitForSelector("tbody tr", new Page.WaitForSelectorOptions().setTimeout(30000));
        page.waitForTimeout(1500);
    }

    private static String getFirstContractName(Page page) throws Exception {
        page.waitForSelector("tbody tr", new Page.WaitForSelectorOptions().setTimeout(30000));
        Locator firstRow = page.locator("tbody tr").first();
        return firstRow.locator("td").nth(2).textContent().trim();
    }

    // 公共方法获取Application list页面第一条数据并点击编辑按钮
    private static void getFirstTableRowAndClickEditButton(Page page) {
        Locator firstRow = page.locator("tbody tr").first();
        Locator editButton = firstRow.locator("td.operation button");
        if (editButton.count() > 0) {
            editButton.click();
            // 等待编辑页面加载
            page.waitForTimeout(3000);
        } else {
            appendLog("Edit button not found");
        }
    }

    // 编辑应用并接受申请
    private static void acceptApplication(Page page) throws Exception {
        // 显式刷新列表
        navigateToApplicationList(page);
        // 获取Application list页面第一条数据点击编辑按钮
        getFirstTableRowAndClickEditButton(page);

        // 点击接受申请按钮
        page.locator("button:has-text('Accept Application')").click();

        // 处理确认弹窗
        handleConfirmationDialog(page);

        // 等待状态更新
        page.waitForSelector(".el-message--success",
                new Page.WaitForSelectorOptions().setTimeout(5000));
        page.waitForTimeout(1500);

    }

    // 编辑应用并发送确认邮件
    private static void issueAcknowledgement(Page page, String[] data) throws Exception {
        navigateToApplicationList(page);

        // 获取Application list页面第一条数据点击编辑按钮
        getFirstTableRowAndClickEditButton(page);

        // 点击发送确认按钮
        appendLog("Click on the Issue Acknowledgement button...");
        page.locator("button:has-text('Issue Acknowledgement')").click();

        // 等待页面跳转到draftDetail页面
        appendLog("Waiting to jump to the draftDetail page...");
        page.waitForURL("**/draftDetail**", new Page.WaitForURLOptions().setTimeout(60000));
        appendLog("Current page URL: " + page.url());

        // 等待新页面元素加载完成
        page.waitForLoadState(LoadState.NETWORKIDLE);

        // 选择Template下拉选择第一个选项
        Locator templateSelect = page.locator("label:has-text('Template:') + div .el-select__wrapper");
        templateSelect.click();
        Locator firstOption = page.locator(".el-select-dropdown__item").first();
        firstOption.click();
        appendLog("Select Template: Drop down and choose the first option to complete");

        // 填充主题字段
        Locator subjectInput = page.locator("label:has-text('Subject:') + div input");
        subjectInput.fill(data[EMAIL_SUBJECT]);
        appendLog("Subject filling completed");
        page.waitForTimeout(1000); //等待填充完成

        // 发送邮件
        page.locator("button:has-text('Send')").click();
        appendLog("Click the send button");

        // 等待状态更新
        page.waitForSelector(".el-message--success",
                new Page.WaitForSelectorOptions().setTimeout(5000));
        page.waitForTimeout(2000);

    }

    // 编辑应用并完成会议
    private static void finishMeeting(Page page) throws Exception {
        navigateToApplicationList(page);

        // 获取Application list页面第一条数据点击编辑按钮
        getFirstTableRowAndClickEditButton(page);

        // 点击完成会议按钮
        page.locator("button:has-text('Finish ID Meeting')").click();

        // 处理确认弹窗
        handleConfirmationDialog(page);

        // 等待状态更新
        page.waitForSelector(".el-message--success",
                new Page.WaitForSelectorOptions().setTimeout(5000));
        page.waitForTimeout(2000);

    }

    // 编辑应用并做决策
    private static void makeDecision(Page page) throws Exception {
        navigateToApplicationList(page);

        // 获取Application list页面第一条数据点击编辑按钮
        getFirstTableRowAndClickEditButton(page);

        // 点击决策按钮
        page.locator("button:has-text('ID Meeting Decision')").click();
        appendLog("The decision button has been clicked");

        // 弹窗等待
        page.waitForSelector(".bd-assessment-dialog:visible",
                new Page.WaitForSelectorOptions().setTimeout(60000));
        appendLog("Decision popup loaded");

        // 下拉框定位器
        Locator dropdown = page.locator(
                "//div[contains(@class, 'bd-assessment-dialog')]" +
                        "//label[contains(., 'Please choose ID Meeting Result')]" +
                        "/following-sibling::div//div[contains(@class, 'el-select__wrapper')]"
        );

        // 确保下拉框可交互
        dropdown.waitFor(new Locator.WaitForOptions()
                .setState(WaitForSelectorState.VISIBLE)
                .setTimeout(15000));
        dropdown.click();
        appendLog("The dropdown menu has been clicked");

        // 使用键盘操作选择第一个选项
        page.keyboard().press("ArrowDown");
        page.waitForTimeout(300);
        page.keyboard().press("Enter");
        appendLog("Select the first option using the keyboard");

        // 点击确认按钮
        page.locator(".el-dialog__footer button:has-text('Confirm')").click();
        appendLog("The confirm button has been clicked");

        // 等待状态更新
        page.waitForSelector(".el-message--success",
                new Page.WaitForSelectorOptions().setTimeout(15000));
        page.waitForTimeout(1500);

    }

    // 编辑应用并请求批准
    private static void requestApproval(Page page) throws Exception {
        // 仍在编辑页面，无需再次打开

        // 点击请求批准按钮
        page.locator("button:has-text('Follow Up / Request for Approval')").click();

        // 等待弹窗出现
        page.waitForSelector(".el-message-box:visible",
                new Page.WaitForSelectorOptions().setTimeout(30000));

        // 点击确认按钮
        page.locator(".el-message-box__btns button.el-button--primary:has-text('Yes')").click();

        // 等待弹窗消失
        page.waitForSelector(".el-message-box",
                new Page.WaitForSelectorOptions().setState(WaitForSelectorState.HIDDEN).setTimeout(30000));

        // 等待状态更新
        page.waitForSelector(".el-message--success",
                new Page.WaitForSelectorOptions().setTimeout(5000));
        page.waitForTimeout(1500);

    }

    // 编辑应用并确认输入
    private static void confirmInput(Page page) throws Exception {
        navigateToApplicationList(page);

        // 获取Application list页面第一条数据点击编辑按钮
        getFirstTableRowAndClickEditButton(page);

        // 点击确认输入按钮
        page.locator("button:has-text('Confirm Input')").click();

        // 处理确认弹窗
        handleConfirmationDialog(page);

        // 等待状态更新
        page.waitForSelector(".el-message--success",
                new Page.WaitForSelectorOptions().setTimeout(5000));
        page.waitForTimeout(1500);

    }

    // 编辑应用并起草通知信
    private static void draftNotification(Page page, String[] data) throws Exception {
        navigateToApplicationList(page);

        // 获取Application list页面第一条数据点击编辑按钮
        getFirstTableRowAndClickEditButton(page);

        // 点击起草通知信按钮
        page.locator("button:has-text('Draft Notification Letter')").click();

        // 等待页面跳转到draftDetail页面
        appendLog("Waiting to jump to the draftDetail page...");
        page.waitForURL("**/draftDetail**", new Page.WaitForURLOptions().setTimeout(60000));
        appendLog("Current page URL: " + page.url());

        // 等待新页面元素加载完成
        page.waitForLoadState(LoadState.NETWORKIDLE);

        // 选择Template下拉选择第一个选项
        Locator templateSelect = page.locator("label:has-text('Template:') + div .el-select__wrapper");
        templateSelect.click();
        Locator firstOption = page.locator(".el-select-dropdown__item").first();
        firstOption.click();
        appendLog("Select Template: Drop down and choose the first option to complete");

        // 填充主题字段
        Locator subjectInput = page.locator("label:has-text('Subject:') + div input");
        subjectInput.fill(data[EMAIL_SUBJECT]);
        appendLog("Subject filling completed");

        // 发送邮件
        page.locator("button:has-text('Send')").click();

        // 等待状态更新
        page.waitForSelector(".el-message--success",
                new Page.WaitForSelectorOptions().setTimeout(5000));

        // 等待操作完成
        page.waitForTimeout(3000);

        // 显式刷新列表
        navigateToApplicationList(page);
    }

    // 处理确认弹窗的通用方法
    private static void handleConfirmationDialog(Page page) throws Exception {
        // 等待弹窗出现
        page.waitForSelector(".el-message-box:visible",
                new Page.WaitForSelectorOptions().setTimeout(30000));

        // 点击确认按钮
        page.locator(".el-message-box__btns button.el-button--primary:has-text('Confirm')").click();

        // 等待弹窗消失
        page.waitForSelector(".el-message-box",
                new Page.WaitForSelectorOptions().setState(WaitForSelectorState.HIDDEN).setTimeout(30000));
    }

    // 进度条更新
    private static void updateProgress(JProgressBar progressBar, int value, String message) {
        SwingUtilities.invokeLater(() -> {
            progressBar.setValue(value);
            progressBar.setString(message);
            logArea.append(message + "\n");
            logArea.setCaretPosition(logArea.getDocument().getLength());
        });
    }

    // 日志输出
    private static void appendLog(String message) {
        SwingUtilities.invokeLater(() -> {
            logArea.append(message + "\n");
            logArea.setCaretPosition(logArea.getDocument().getLength());
        });
    }

    private static List<String[]> readCSV(String csvPath) {
        List<String[]> data = new ArrayList<>();
        try (BufferedReader br = new BufferedReader(new FileReader(csvPath))) {
            br.readLine(); // 跳过标题行
            String line;
            while ((line = br.readLine()) != null) {
                data.add(line.split(","));
            }
            appendLog("Successfully read CSV file: " + csvPath + ",   Number of records: " + data.size());
        } catch (IOException e) {
            appendLog("Error reading CSV: " + e.getMessage());
        }
        return data;
    }

    private static void login(Page page, String url) throws Exception {
        page.navigate(url);

        // 确保页面加载完成
        page.waitForSelector(".el-select__wrapper");

        // 填写登录表单
        page.locator(".el-select__wrapper").first().click();
        page.getByText("demo_system_admin").click();
        page.locator("div:nth-child(4) > .el-form-item__content > .el-select > .el-select__wrapper").click();
        page.getByRole(AriaRole.OPTION, new Page.GetByRoleOptions().setName("DEVB")).click();
        page.getByRole(AriaRole.BUTTON, new Page.GetByRoleOptions().setName("Login")).click();

        // 显式等待登录成功（例如：检查导航到首页后的元素）
        page.waitForSelector("text=Applications", new Page.WaitForSelectorOptions().setTimeout(60000));
    }

    private static void createApplication(Page page, String[] data) throws Exception {
        // 导航到应用创建
        page.locator(".el-menu-item:has-text('Applications')").click();
        page.waitForSelector("button:has-text('Add Application')", new Page.WaitForSelectorOptions().setTimeout(30000));
        page.getByRole(AriaRole.BUTTON, new Page.GetByRoleOptions().setName("Add Application")).click();

        // 填写基本信息
        page.getByRole(AriaRole.TEXTBOX, new Page.GetByRoleOptions().setName("Name of principal contractor(Eng)"))
                .fill(data[PRINCIPAL_NAME]);

        fillField(page, "B.R. certificate no", data[B_R_CERTIFICATE_NO]);
        fillField(page, "* Address:", data[ADDRESS]);
        fillField(page, "* Name:", data[CONTACT_NAME]);
        fillField(page, "* Post Title:", data[POST_TITLE]);
        fillField(page, "* Tel No.:", data[TEL_NO]);
        fillField(page, "* Email::", data[EMAIL]);
        fillField(page, "* Contract no.:", data[CONTRACT_NO]);
        fillField(page, "* Name of the contract (Eng):", data[CONTRACT_NAME]);
        fillField(page, "* Work site:", data[WORK_SITE]);
        fillField(page, "* Commencement date:", data[COMMENCEMENT_DATE]);
        fillField(page, "* Target completion date:", data[COMPLETION_DATE]);

        // 单选按钮选项（根据实际需求调整）
        page.getByText("For private sector works").click();
        page.getByText("Entry requirements are").click();
        page.getByText("Applicant undertaken the work").click();

        // 下拉选择
        selectDropdown(page, "Type of accommodation", "All imported labour living in");
        selectDropdown(page, "Address", "site in part 3.3");

        // 继续填写
        fillField(page, "* Specific implementation", data[IMPLEMENTATION_MEASURES]);
        page.getByText("Applicant undertaken above").click();

        // 第一次保存
        page.getByRole(AriaRole.BUTTON, new Page.GetByRoleOptions().setName("Save")).click();

        // 根据页面要求第二次保存
        page.getByRole(AriaRole.BUTTON, new Page.GetByRoleOptions().setName("Save")).click();

        // 填写人力计划
        page.getByRole(AriaRole.TAB, new Page.GetByRoleOptions().setName("Manpower Plan")).click();
        page.locator("div").filter(new Locator.FilterOptions().setHasText(Pattern.compile("^Part 2a: Skilled WorkerAdd$")))
                .getByRole(AriaRole.BUTTON).click();

        // 选择工种
        selectDropdown(page, "Trade", data[TRADE]);
        page.getByLabel("Manpower Requirement By Trade").getByRole(AriaRole.BUTTON, new Locator.GetByRoleOptions().setName("Add")).click();

        // 填写第一行数据
        fillDateRangeField(page, 1, data[PERIOD_START], data[PERIOD_END]);
        fillManpowerField(page, 1, data[MANPOWER_REQUIRED]);
        fillField(page, "(B2) No. of Local Labour:", data[LOCAL_LABOUR]);
        fillField(page, "(B3a.1) No. of Imported Labour(Applied):", data[IMPORTED_LABOUR]);
        fillField(page, "(B3b) Total No. of Month:", data[TOTAL_MONTHS]);
        fillField(page, "(B3c) Monthly Wage ($):", data[MONTHLY_WAGE]);

        // 人力计划保存
        page.getByLabel("Manpower Requirement By Trade").getByRole(AriaRole.BUTTON, new Locator.GetByRoleOptions().setName("Save")).click();

        // 填写总结数据
        fillField(page, "Number of skilled workers and technicians:", data[SKILLED_WORKERS]);
        fillField(page, "Number of others(excluding managerial staff):", data[OTHERS]);
        fillField(page, "Number of Imported labour applied for:", data[IMPORTED_APPLIED]);

        // 最后保存
        page.locator("#app").getByRole(AriaRole.BUTTON, new Locator.GetByRoleOptions().setName("Save")).click();

        // 等待操作完成
        page.waitForTimeout(1000);
    }

    // 通用字段填写
    private static void fillField(Page page, String fieldName, String value) throws Exception {
        appendLog("Attempt to fill in fields: " + fieldName);

        // 1. 尝试普通文本输入框
        Locator textbox = page.getByRole(AriaRole.TEXTBOX,
                new Page.GetByRoleOptions().setName(fieldName));

        if (textbox.count() > 0) {
            textbox.fill(value);
            appendLog("  Successfully filled with TEXTBOX positioning");
            return;
        }

        // 2. 尝试数字输入框 (SPINBUTTON)
        Locator spinbutton = page.getByRole(AriaRole.SPINBUTTON,
                new Page.GetByRoleOptions().setName(fieldName));

        if (spinbutton.count() > 0) {
            spinbutton.fill(value);
            appendLog("  Successfully filled using SPINBUTTON positioning");
            return;
        }

        // 3. 尝试通用输入框定位
        String xpath = String.format(
                "//*[contains(text(), '%s')]/ancestor::div[contains(@class, 'el-form-item')]//input",
                fieldName
        );

        Locator genericInput = page.locator(xpath);
        if (genericInput.count() > 0) {
            genericInput.fill(value);
            appendLog("Successfully filled using universal XPATH positioning");
            return;
        }

        // 4. 最后尝试：聚焦后输入
        try {
            page.locator(xpath).focus();
            page.keyboard().type(value);
            appendLog("  Successfully entered using keyboard");
        } catch (Exception e) {
            appendLog("Unable to locate field: " + fieldName);
            throw e;
        }
    }

    // 下拉选择
    private static void selectDropdown(Page page, String dropdownName, String option) {
        page.locator("div").filter(new Locator.FilterOptions().setHasText(Pattern.compile("^" + dropdownName + "$")))
                .nth(2).click();
        page.getByRole(AriaRole.OPTION, new Page.GetByRoleOptions().setName(option)).click();
    }

    // 填充日期范围字段
    private static void fillDateRangeField(Page page, int rowIndex, String startValue, String endValue) throws Exception {
        // 定位指定行的日期选择器
        Locator dateRangePicker = page.locator(
                String.format("//tbody/tr[%d]/td[1]//div[contains(@class, 'el-date-editor')]", rowIndex)
        );

        // 点击打开日期选择器
        dateRangePicker.click();

        // 分别填充开始和结束月份
        Locator startInput = dateRangePicker.locator("input").first();
        startInput.fill(startValue);

        Locator endInput = dateRangePicker.locator("input").nth(1);
        endInput.fill(endValue);

        // 按Enter确认
        dateRangePicker.press("Enter");
    }

    private static void fillManpowerField(Page page, int rowIndex, String value) {
        // 定位指定行的人力需求输入框
        Locator input = page.locator(
                String.format("//tbody/tr[%d]/td[3]//input", rowIndex)
        );
        input.fill(value);
    }
}