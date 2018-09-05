package bean;

/**
 * @author wangganglong
 * @date 2018/09/04 上午10:54
 **/
public class StatisticalBean {
    private String statisticalName;
    private int statisticalNum;

    public StatisticalBean() {
    }

    public StatisticalBean(String statisticalName, int statisticalNum) {
        this.statisticalName = statisticalName;
        this.statisticalNum = statisticalNum;
    }

    public String getStatisticalName() {
        return statisticalName;
    }

    public void setStatisticalName(String statisticalName) {
        this.statisticalName = statisticalName;
    }

    public int getStatisticalNum() {
        return statisticalNum;
    }

    public void setStatisticalNum(int statisticalNum) {
        this.statisticalNum = statisticalNum;
    }
}
