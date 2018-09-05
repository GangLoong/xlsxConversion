import bean.StatisticalBean;
import bean.XLSLBean;
import org.apache.commons.beanutils.BeanUtils;
import org.apache.commons.beanutils.BeanUtilsBean;
import org.apache.commons.collections.CollectionUtils;

import java.beans.PropertyDescriptor;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;


public class ExcelUtils<T> {

    List<T> list = new ArrayList<T>();

    public static void main(String[] args) {
        String suffixName = ".xlsx";
        String tableName = "2016.1";
        String url = "/Users/wangganglong/Projects/ProcessingComplete/" + tableName + suffixName;
        File xls = new File(url);
        Map<String, String> headMap = new HashMap<String, String>();
        headMap.put("案件状态", "caseState");
        headMap.put("任务号", "taskId");
        headMap.put("案件来源", "caseSource");
        headMap.put("发现时间", "findTime");
        headMap.put("流转方向", "flowDirection");
        headMap.put("案件属性", "caseAttributes");
        headMap.put("案件大类", "casesBigClass");
        headMap.put("案件小类", "casesSmallClass");
        headMap.put("案件子类", "casesSonClass");
        headMap.put("街镇", "streetTown");
        headMap.put("发生地址", "address");
        headMap.put("立案部门", "initiateDepartment");
        headMap.put("派遣时间", "dispatchTime");
        headMap.put("主责部门", "mainResponsibleDepartments");
        headMap.put("三级主责部门", "thirdLevelMainResponsibleDepartments");
        headMap.put("结案时间", "closingTime");
        System.out.println("开始");
        try {
            InputStream inputStream = new FileInputStream(xls);
            ExcelUtils<XLSLBean> eu = new ExcelUtils<XLSLBean>();
            List<XLSLBean> list = eu.readBean(inputStream, XLSLBean.class, headMap);
            //createTable(tableName);
            statistical(list, tableName);
            System.out.println(list.size());
            System.out.println("结束");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static String createTable(String tableName) {
        String sql = "CREATE TABLE `" + tableName + "` (" +
                "  `id` int(11) NOT NULL AUTO_INCREMENT COMMENT 'Id'," +
                "  `statisticalName` varchar(225) NOT NULL COMMENT '统计名称'," +
                "  `statisticalNum` int(11) NOT NULL COMMENT '统计数量'," +
                "  `createDate` datetime NOT NULL COMMENT '创建时间'," +
                "  PRIMARY KEY (`id`)" +
                ") ENGINE=InnoDB AUTO_INCREMENT=1 DEFAULT CHARSET=utf8;";
        mysql(sql);
        System.out.println("建表成功！");
        return "ok";
    }

    private static List statistical(List<XLSLBean> xlslBeans, String tableName) throws ParseException {
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy.MM.dd HH:mm:ss");
        Date date = sdf.parse(tableName + ".1 00:00:00");
        SimpleDateFormat sd = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        //获取String类型的时间
        String createdate = sd.format(date);
        List<StatisticalBean> statisticalBeans = new ArrayList<StatisticalBean>();
        if (CollectionUtils.isEmpty(xlslBeans)) {
            return statisticalBeans;
        }
        int a = 0;
        int b = 0;
        int c = 0;
        int d = 0;
        int e = 0;
        int f = 0;
        int g = 0;
        int h = 0;
        int i = 0;
        int j = 0;
        int k = 0;
        int k_a = 0;
        int k_b = 0;
        int k_c = 0;
        int k_d = 0;
        int k_e = 0;
        int l = 0;
        int l_a = 0;
        int l_b = 0;
        int l_c = 0;
        int l_d = 0;
        int l_e = 0;
        int m = 0;
        int m_a = 0;
        int m_b = 0;
        int m_c = 0;
        int m_d = 0;
        int m_e = 0;
        int n = 0;
        int n_a = 0;
        int n_b = 0;
        int n_c = 0;
        int n_d = 0;
        int n_e = 0;
        int o = 0;
        int o_a = 0;
        int o_b = 0;
        int o_c = 0;
        int o_d = 0;
        int o_e = 0;
        int p = 0;
        int p_a = 0;
        int p_b = 0;
        int p_c = 0;
        int p_d = 0;
        int p_e = 0;
        int q = 0;
        int q_a = 0;
        int q_b = 0;
        int q_c = 0;
        int q_d = 0;
        int q_e = 0;
        int r = 0;
        int r_a = 0;
        int r_b = 0;
        int r_c = 0;
        int r_d = 0;
        int r_e = 0;
        int s = 0;
        int s_a = 0;
        int s_b = 0;
        int s_c = 0;
        int s_d = 0;
        int s_e = 0;
        int t = 0;
        int u = 0;
        int u_a = 0;
        int u_b = 0;
        int u_c = 0;
        int u_d = 0;
        int u_e = 0;
        int v = 0;
        int w = 0;
        for (XLSLBean xlslBean :
                xlslBeans) {
            if ("暴露垃圾".equals(xlslBean.getCasesSmallClass())) {
                a++;
            }
            if ("乱涂写、乱张贴、乱刻画".equals(xlslBean.getCasesSmallClass())) {
                b++;
            }
            if ("占道无证照经营".equals(xlslBean.getCasesSmallClass())) {
                c++;
            }
            if ("跨门营业".equals(xlslBean.getCasesSmallClass())) {
                d++;
            }
            if ("擅自占用道路堆物、施工".equals(xlslBean.getCasesSmallClass())) {
                e++;
            }

            if ("居委采集".equals(xlslBean.getCaseSource())) {
                f++;
            }
            if ("视频上报".equals(xlslBean.getCaseSource())) {
                g++;
            }
            if ("监督员上报".equals(xlslBean.getCaseSource())) {
                h++;
            }
            if ("工作站采集".equals(xlslBean.getCaseSource())) {
                i++;
            }
            if ("区级督查".equals(xlslBean.getCaseSource())) {
                j++;
            }
            if ("12319上报".equals(xlslBean.getCaseSource()) || "12345(12319转派)".equals(xlslBean.getCaseSource())) {
                v++;
            }
            if ("区级督查".equals(xlslBean.getCaseSource())) {
                w++;
            }

            if ("南京东路街道".equals(xlslBean.getStreetTown())) {
                k++;
                if ("暴露垃圾".equals(xlslBean.getCasesSmallClass())) {
                    k_a++;
                }
                if ("乱涂写、乱张贴、乱刻画".equals(xlslBean.getCasesSmallClass())) {
                    k_b++;
                }
                if ("占道无证照经营".equals(xlslBean.getCasesSmallClass())) {
                    k_c++;
                }
                if ("跨门营业".equals(xlslBean.getCasesSmallClass())) {
                    k_d++;
                }
                if ("擅自占用道路堆物、施工".equals(xlslBean.getCasesSmallClass())) {
                    k_e++;
                }
            }
            if ("外滩街道".equals(xlslBean.getStreetTown())) {
                l++;
                if ("暴露垃圾".equals(xlslBean.getCasesSmallClass())) {
                    l_a++;
                }
                if ("乱涂写、乱张贴、乱刻画".equals(xlslBean.getCasesSmallClass())) {
                    l_b++;
                }
                if ("占道无证照经营".equals(xlslBean.getCasesSmallClass())) {
                    l_c++;
                }
                if ("跨门营业".equals(xlslBean.getCasesSmallClass())) {
                    l_d++;
                }
                if ("擅自占用道路堆物、施工".equals(xlslBean.getCasesSmallClass())) {
                    l_e++;
                }
            }
            if ("瑞金二路街道".equals(xlslBean.getStreetTown())) {
                m++;
                if ("暴露垃圾".equals(xlslBean.getCasesSmallClass())) {
                    m_a++;
                }
                if ("乱涂写、乱张贴、乱刻画".equals(xlslBean.getCasesSmallClass())) {
                    m_b++;
                }
                if ("占道无证照经营".equals(xlslBean.getCasesSmallClass())) {
                    m_c++;
                }
                if ("跨门营业".equals(xlslBean.getCasesSmallClass())) {
                    m_d++;
                }
                if ("擅自占用道路堆物、施工".equals(xlslBean.getCasesSmallClass())) {
                    m_e++;
                }
            }
            if ("淮海中路街道".equals(xlslBean.getStreetTown())) {
                u++;
                if ("暴露垃圾".equals(xlslBean.getCasesSmallClass())) {
                    u_a++;
                }
                if ("乱涂写、乱张贴、乱刻画".equals(xlslBean.getCasesSmallClass())) {
                    u_b++;
                }
                if ("占道无证照经营".equals(xlslBean.getCasesSmallClass())) {
                    u_c++;
                }
                if ("跨门营业".equals(xlslBean.getCasesSmallClass())) {
                    u_d++;
                }
                if ("擅自占用道路堆物、施工".equals(xlslBean.getCasesSmallClass())) {
                    u_e++;
                }
            }
            if ("豫园街道".equals(xlslBean.getStreetTown())) {
                n++;
                if ("暴露垃圾".equals(xlslBean.getCasesSmallClass())) {
                    n_a++;
                }
                if ("乱涂写、乱张贴、乱刻画".equals(xlslBean.getCasesSmallClass())) {
                    n_b++;
                }
                if ("占道无证照经营".equals(xlslBean.getCasesSmallClass())) {
                    n_c++;
                }
                if ("跨门营业".equals(xlslBean.getCasesSmallClass())) {
                    n_d++;
                }
                if ("擅自占用道路堆物、施工".equals(xlslBean.getCasesSmallClass())) {
                    n_e++;
                }
            }
            if ("打浦桥街道".equals(xlslBean.getStreetTown())) {
                o++;
                if ("暴露垃圾".equals(xlslBean.getCasesSmallClass())) {
                    o_a++;
                }
                if ("乱涂写、乱张贴、乱刻画".equals(xlslBean.getCasesSmallClass())) {
                    o_b++;
                }
                if ("占道无证照经营".equals(xlslBean.getCasesSmallClass())) {
                    o_c++;
                }
                if ("跨门营业".equals(xlslBean.getCasesSmallClass())) {
                    o_d++;
                }
                if ("擅自占用道路堆物、施工".equals(xlslBean.getCasesSmallClass())) {
                    o_e++;
                }
            }
            if ("老西门街道".equals(xlslBean.getStreetTown())) {
                p++;
                if ("暴露垃圾".equals(xlslBean.getCasesSmallClass())) {
                    p_a++;
                }
                if ("乱涂写、乱张贴、乱刻画".equals(xlslBean.getCasesSmallClass())) {
                    p_b++;
                }
                if ("占道无证照经营".equals(xlslBean.getCasesSmallClass())) {
                    p_c++;
                }
                if ("跨门营业".equals(xlslBean.getCasesSmallClass())) {
                    p_d++;
                }
                if ("擅自占用道路堆物、施工".equals(xlslBean.getCasesSmallClass())) {
                    p_e++;
                }
            }
            if ("小东门街道".equals(xlslBean.getStreetTown())) {
                q++;
                if ("暴露垃圾".equals(xlslBean.getCasesSmallClass())) {
                    q_a++;
                }
                if ("乱涂写、乱张贴、乱刻画".equals(xlslBean.getCasesSmallClass())) {
                    q_b++;
                }
                if ("占道无证照经营".equals(xlslBean.getCasesSmallClass())) {
                    q_c++;
                }
                if ("跨门营业".equals(xlslBean.getCasesSmallClass())) {
                    q_d++;
                }
                if ("擅自占用道路堆物、施工".equals(xlslBean.getCasesSmallClass())) {
                    q_e++;
                }
            }
            if ("五里桥街道".equals(xlslBean.getStreetTown())) {
                r++;
                if ("暴露垃圾".equals(xlslBean.getCasesSmallClass())) {
                    r_a++;
                }
                if ("乱涂写、乱张贴、乱刻画".equals(xlslBean.getCasesSmallClass())) {
                    r_b++;
                }
                if ("占道无证照经营".equals(xlslBean.getCasesSmallClass())) {
                    r_c++;
                }
                if ("跨门营业".equals(xlslBean.getCasesSmallClass())) {
                    r_d++;
                }
                if ("擅自占用道路堆物、施工".equals(xlslBean.getCasesSmallClass())) {
                    r_e++;
                }
            }
            if ("半淞园街道".equals(xlslBean.getStreetTown())) {
                s++;
                if ("暴露垃圾".equals(xlslBean.getCasesSmallClass())) {
                    s_a++;
                }
                if ("乱涂写、乱张贴、乱刻画".equals(xlslBean.getCasesSmallClass())) {
                    s_b++;
                }
                if ("占道无证照经营".equals(xlslBean.getCasesSmallClass())) {
                    s_c++;
                }
                if ("跨门营业".equals(xlslBean.getCasesSmallClass())) {
                    s_d++;
                }
                if ("擅自占用道路堆物、施工".equals(xlslBean.getCasesSmallClass())) {
                    s_e++;
                }
            }

            if ("已结案".equals(xlslBean.getCaseState())) {
                t++;
            }
        }
        statisticalBeans.add(new StatisticalBean("暴露垃圾", a));
        statisticalBeans.add(new StatisticalBean("乱涂写、乱张贴、乱刻画", b));
        statisticalBeans.add(new StatisticalBean("占道无证照经营", c));
        statisticalBeans.add(new StatisticalBean("跨门营业", d));
        statisticalBeans.add(new StatisticalBean("擅自占用道路堆物、施工", e));
        statisticalBeans.add(new StatisticalBean("居委采集", f));
        statisticalBeans.add(new StatisticalBean("视频上报", g));
        statisticalBeans.add(new StatisticalBean("监督员上报", h));
        statisticalBeans.add(new StatisticalBean("工作站采集", i));
        statisticalBeans.add(new StatisticalBean("区级督查", j));
        statisticalBeans.add(new StatisticalBean("案件来源（12319+12345)", v));
        statisticalBeans.add(new StatisticalBean("网络", w));
        statisticalBeans.add(new StatisticalBean("其他", xlslBeans.size() - w + v));

        statisticalBeans.add(new StatisticalBean("南京东路街道", k));
        statisticalBeans.add(new StatisticalBean("南京东路街道-暴露垃圾", k_a));
        statisticalBeans.add(new StatisticalBean("南京东路街道-乱涂写、乱张贴、乱刻画", k_b));
        statisticalBeans.add(new StatisticalBean("南京东路街道-占道无证照经营", k_c));
        statisticalBeans.add(new StatisticalBean("南京东路街道-跨门营业", k_d));
        statisticalBeans.add(new StatisticalBean("南京东路街道-擅自占用道路堆物、施工", k_e));

        statisticalBeans.add(new StatisticalBean("外滩街道", l));
        statisticalBeans.add(new StatisticalBean("外滩街道-暴露垃圾", l_a));
        statisticalBeans.add(new StatisticalBean("外滩街道-乱涂写、乱张贴、乱刻画", l_b));
        statisticalBeans.add(new StatisticalBean("外滩街道-占道无证照经营", l_c));
        statisticalBeans.add(new StatisticalBean("外滩街道-跨门营业", l_d));
        statisticalBeans.add(new StatisticalBean("外滩街道-擅自占用道路堆物、施工", l_e));

        statisticalBeans.add(new StatisticalBean("瑞金二路街道", m));
        statisticalBeans.add(new StatisticalBean("瑞金二路街道-暴露垃圾", m_a));
        statisticalBeans.add(new StatisticalBean("瑞金二路街道-乱涂写、乱张贴、乱刻画", m_b));
        statisticalBeans.add(new StatisticalBean("瑞金二路街道-占道无证照经营", m_c));
        statisticalBeans.add(new StatisticalBean("瑞金二路街道-跨门营业", m_d));
        statisticalBeans.add(new StatisticalBean("瑞金二路街道-擅自占用道路堆物、施工", m_e));

        statisticalBeans.add(new StatisticalBean("淮海中路街道", u));
        statisticalBeans.add(new StatisticalBean("淮海中路街道-暴露垃圾", u_a));
        statisticalBeans.add(new StatisticalBean("淮海中路街道-乱涂写、乱张贴、乱刻画", u_b));
        statisticalBeans.add(new StatisticalBean("淮海中路街道-占道无证照经营", u_c));
        statisticalBeans.add(new StatisticalBean("淮海中路街道-跨门营业", u_d));
        statisticalBeans.add(new StatisticalBean("淮海中路街道-擅自占用道路堆物、施工", u_e));

        statisticalBeans.add(new StatisticalBean("豫园街道", n));
        statisticalBeans.add(new StatisticalBean("豫园街道-暴露垃圾", n_a));
        statisticalBeans.add(new StatisticalBean("豫园街道-乱涂写、乱张贴、乱刻画", n_b));
        statisticalBeans.add(new StatisticalBean("豫园街道-占道无证照经营", n_c));
        statisticalBeans.add(new StatisticalBean("豫园街道-跨门营业", n_d));
        statisticalBeans.add(new StatisticalBean("豫园街道-擅自占用道路堆物、施工", n_e));

        statisticalBeans.add(new StatisticalBean("打浦桥街道", o));
        statisticalBeans.add(new StatisticalBean("打浦桥街道-暴露垃圾", o_a));
        statisticalBeans.add(new StatisticalBean("打浦桥街道-乱涂写、乱张贴、乱刻画", o_b));
        statisticalBeans.add(new StatisticalBean("打浦桥街道-占道无证照经营", o_c));
        statisticalBeans.add(new StatisticalBean("打浦桥街道-跨门营业", o_d));
        statisticalBeans.add(new StatisticalBean("打浦桥街道-擅自占用道路堆物、施工", o_e));

        statisticalBeans.add(new StatisticalBean("老西门街道", p));
        statisticalBeans.add(new StatisticalBean("老西门街道-暴露垃圾", p_a));
        statisticalBeans.add(new StatisticalBean("老西门街道-乱涂写、乱张贴、乱刻画", p_b));
        statisticalBeans.add(new StatisticalBean("老西门街道-占道无证照经营", p_c));
        statisticalBeans.add(new StatisticalBean("老西门街道-跨门营业", p_d));
        statisticalBeans.add(new StatisticalBean("老西门街道-擅自占用道路堆物、施工", p_e));

        statisticalBeans.add(new StatisticalBean("小东门街道", q));
        statisticalBeans.add(new StatisticalBean("小东门街道-暴露垃圾", q_a));
        statisticalBeans.add(new StatisticalBean("小东门街道-乱涂写、乱张贴、乱刻画", q_b));
        statisticalBeans.add(new StatisticalBean("小东门街道-占道无证照经营", q_c));
        statisticalBeans.add(new StatisticalBean("小东门街道-跨门营业", q_d));
        statisticalBeans.add(new StatisticalBean("小东门街道-擅自占用道路堆物、施工", q_e));

        statisticalBeans.add(new StatisticalBean("五里桥街道", r));
        statisticalBeans.add(new StatisticalBean("五里桥街道-暴露垃圾", r_a));
        statisticalBeans.add(new StatisticalBean("五里桥街道-乱涂写、乱张贴、乱刻画", r_b));
        statisticalBeans.add(new StatisticalBean("五里桥街道-占道无证照经营", r_c));
        statisticalBeans.add(new StatisticalBean("五里桥街道-跨门营业", r_d));
        statisticalBeans.add(new StatisticalBean("五里桥街道-擅自占用道路堆物、施工", r_e));

        statisticalBeans.add(new StatisticalBean("半淞园街道", s));
        statisticalBeans.add(new StatisticalBean("半淞园街道-暴露垃圾", s_a));
        statisticalBeans.add(new StatisticalBean("半淞园街道-乱涂写、乱张贴、乱刻画", s_b));
        statisticalBeans.add(new StatisticalBean("半淞园街道-占道无证照经营", s_c));
        statisticalBeans.add(new StatisticalBean("半淞园街道-跨门营业", s_d));
        statisticalBeans.add(new StatisticalBean("半淞园街道-擅自占用道路堆物、施工", s_e));

        statisticalBeans.add(new StatisticalBean("已结案", t));
        statisticalBeans.add(new StatisticalBean("处理中", xlslBeans.size() - t));
        statisticalBeans.add(new StatisticalBean("受理数", xlslBeans.size()));
        String sql = "INSERT INTO `statistical`.`t_huangpu_statistical`(`statisticalName`, `statisticalNum`,`time`) VALUES ";
        StringBuilder valueSql = new StringBuilder();
        for (int fi = 0; fi < statisticalBeans.size(); fi++) {
            valueSql.append("('").append(statisticalBeans.get(fi).getStatisticalName()).append("',").append(statisticalBeans.get(fi).getStatisticalNum()).append(",'").append(createdate).append("')");
            if (fi != statisticalBeans.size() - 1) {
                valueSql.append(",");
            }
        }
        sql += valueSql;
        mysql(sql);
        return statisticalBeans;
    }

    private static void mysql(String sql) {
        // 驱动程序名
        String driver = "com.mysql.jdbc.Driver";

        // URL指向要访问的数据库名scutcs
        String url = "jdbc:mysql://127.0.0.1:3306/statistical?useUnicode=true&characterEncoding=utf-8&relaxAutoCommit=true&zeroDateTimeBehavior=convertToNull&allowMultiQueries=true";

        // MySQL配置时的用户名
        String user = "root";

        // MySQL配置时的密码
        String password = "root";
        try {
            // 加载驱动程序
            Class.forName(driver);

            // 连续数据库
            Connection conn = DriverManager.getConnection(url, user, password);

            if (!conn.isClosed()) {
                System.out.println("Succeeded connecting to the Database!");
            }

            // statement用来执行SQL语句
            Statement statement = conn.createStatement();

            // 结果集
            statement.executeUpdate(sql);
            conn.close();
        } catch (ClassNotFoundException e) {
            e.printStackTrace();
        } catch (SQLException e) {
            e.printStackTrace();
        }
    }

    /**
     * 读取Excel2003-2007转换为实体Bean
     *
     * @param inputStream
     * @param tClass       实体类
     * @param headerMapper
     * @return
     * @throws Exception
     */
    public List<T> readBean(InputStream inputStream, Class<T> tClass, Map<String, String> headerMapper)
            throws Exception {
        ArrayList<ArrayList<String>> rows = ExcelIO.readExcel(inputStream, 0);
        for (int k = 1; k < rows.size(); k++) {
            T t = tClass.newInstance();
            for (int num = 0; num < rows.get(0).size(); num++) {
                for (String header : headerMapper.keySet()) {
                    String name = rows.get(0).get(num);
                    if (name.equals(header)) {
                        PropertyDescriptor propertyDescriptor = BeanUtilsBean.getInstance().getPropertyUtils()
                                .getPropertyDescriptor(t, headerMapper.get(header).toString());
                        if (propertyDescriptor != null) {
                            BeanUtils.setProperty(t, headerMapper.get(header).toString(),
                                    ExcelIO.changeType(rows.get(k).get(num), propertyDescriptor.getPropertyType()));
                            break;
                        }
                    }
                }
            }
            list.add(t);
        }
        return list;
    }

}
