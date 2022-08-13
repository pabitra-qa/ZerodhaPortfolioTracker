package ZerodhaPortfolioInsight;

import java.io.FileInputStream;
import java.util.Properties;

public class PropertiesReader {
    private static Properties prop;

    /**
     * This method is used to load the properties from config.properties file
     *
     * @return it returns Properties value as per yje PropertyName passed as parameter
     */
    public static String getPropertyValue(String propertyName){

        prop = new Properties();
        try {
            FileInputStream ip = new FileInputStream("zerodha.properties");
            prop.load(ip);
        }catch(Exception e)
        {

        }
        return prop.getProperty(propertyName);
    }
}
