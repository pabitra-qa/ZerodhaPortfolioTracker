package ZerodhaPortfolioInsight;

import javax.swing.*;
import java.awt.*;
import java.io.IOException;

public class PortfolioScraper {
    private JPanel MainPanel;
    private JTextField usernameTxt;
    private JPasswordField passwordTxt;
    private JPasswordField pinTxt;
    private JButton start;
    private JTextArea opTextArea;
    private JButton emailButton;
    private JButton exitBtn;
    private JTextField emailAddress;
    private JPasswordField emailAppPassword;
    private JButton writeToExcelBtn;
    public JPanel bottomPanel;

    public PortfolioScraper(String username, String password, String pin, String coinHoldings, String email, String appPass) {

        String coin;
        if (username == null && password == null && pin == null && coinHoldings == null && email == null && appPass == null) {
            System.out.println("Please make sure the 'zerodha.properties' exits in the project directory...");
            opTextArea.append("Please make sure the 'zerodha.properties' exits in the project directory...\n");
            opTextArea.append("You can still go ahead by entering the details manually... We will assume you don't have coin holdings...\n");
        }

        if (username != null)
            usernameTxt.setText(username);
        if (password != null)
            passwordTxt.setText(password);
        if (pin != null)
            pinTxt.setText(pin);
        if (email != null)
            emailAddress.setText(email);
        if (appPass != null)
            emailAppPassword.setText(appPass);
        if (coinHoldings != null)
            coin = coinHoldings;
        else
            coin = "no";

        bottomPanel.setVisible(false);

        start.addActionListener(e -> {
            try {
                PortfolioScraperZerodha.textArea = opTextArea;
                if (coin.equalsIgnoreCase("no")) {
                    PortfolioScraperZerodha.mainMethod(usernameTxt.getText(), passwordTxt.getPassword(), pinTxt.getPassword(), "N");
                } else {
                    PortfolioScraperZerodha.mainMethod(usernameTxt.getText(), passwordTxt.getPassword(), pinTxt.getPassword(), "Y");
                }
                bottomPanel.setVisible(true);
            } catch (Exception ex) {
                opTextArea.append(ex.getMessage());
            }
        });
        emailButton.addActionListener(e -> {
            try {
                PortfolioScraperZerodha.sendEmail(emailAddress.getText(), emailAppPassword.getPassword());
            } catch (Exception ex) {
                opTextArea.append(ex.getMessage());
            }
        });
        writeToExcelBtn.addActionListener(e -> {
            try {
                PortfolioScraperZerodha.writeToExcel();
            } catch (IOException ex) {
                opTextArea.append(ex.getMessage());
            }
        });
        exitBtn.addActionListener(e -> System.exit(0));
    }

    public static void main(String[] args) {
        String username = PropertiesReader.getPropertyValue("username");
        String password = PropertiesReader.getPropertyValue("password");
        String pin = PropertiesReader.getPropertyValue("pin");
        String coinHoldings = PropertiesReader.getPropertyValue("coinHoldings");
        String email = PropertiesReader.getPropertyValue("email");
        String appPass = PropertiesReader.getPropertyValue("appPassword");

        PortfolioScraper pf = new PortfolioScraper(username, password, pin, coinHoldings, email, appPass);

        Font font = new Font("Monospaced", Font.PLAIN, 14);
        pf.opTextArea.setFont(font);

        JFrame frame = new JFrame();
        frame.setTitle("Zerodha Portfolio Tracker");
        pf.MainPanel.setSize(new Dimension(1200, 900));
        frame.add(pf.MainPanel);
        frame.setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
        frame.setPreferredSize(new Dimension(1200, 900));
        frame.pack();
        frame.setVisible(true);
        PortfolioScraperZerodha.frame = frame;
    }
}
