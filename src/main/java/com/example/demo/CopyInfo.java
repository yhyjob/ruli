package com.example.demo;

public class CopyInfo {

    private String packnumcek = "";
    private String fundid = "";
    private String ruliobject = "";
    private String pageno;
    private String descsheetname = "";
    private String copycount = "";
    private String fundtitle = "";
    private String srcfileno = "";
    private String baoyou = "";
    private String quanti = "";
    private int beginrow = 0;
    private int endrow = 0;
    private final int sheetindex = 0;
    private final int titleAtColindex = 0;
    private int realcopycount = 0;

    public int getSheetindex() {
        int intdex = 0;
        if (Integer.parseInt(descsheetname.substring(0, 2)) > 10) {
            intdex = Integer.parseInt(descsheetname.substring(0, 2)) - 1;
        } else {
            intdex = Integer.parseInt(descsheetname.substring(0, 2));
        }
        return intdex;
    }

    public int getTitleAtColindex() {
        int index = 0;
        if (getSheetindex() == 1 || getSheetindex() == 2 || getSheetindex() == 5 || getSheetindex() == 7) {
            index = 9;
        } else if (getSheetindex() == 3) {
            index = 12;
        } else if (getSheetindex() == 4 || getSheetindex() == 9 || getSheetindex() == 13 || getSheetindex() == 14) {
            index = 8;
        } else if (getSheetindex() == 8) {
            index = 7;
        } else if (getSheetindex() == 12) {
            index = 6;
        } else if (getSheetindex() == 17) {
            index = 2;
        } else {

        }
        return index;
    }


    public String getPacknumcek() {
        return packnumcek;
    }

    public void setPacknumcek(String packnumcek) {
        this.packnumcek = packnumcek;
    }

    public String getFundid() {
        return fundid;
    }

    public void setFundid(String fundid) {
        this.fundid = fundid;
    }

    public String getRuliobject() {
        return ruliobject;
    }

    public void setRuliobject(String ruliobject) {
        this.ruliobject = ruliobject;
    }

    public String getPageno() {
        return pageno;
    }

    public void setPageno(String pageno) {
        this.pageno = pageno;
    }

    public String getDescsheetname() {
        return descsheetname;
    }

    public void setDescsheetname(String descsheetname) {
        this.descsheetname = descsheetname;
    }

    public String getCopycount() {
        return copycount;
    }

    public void setCopycount(String copycount) {
        this.copycount = copycount;
    }

    public String getFundtitle() {
        return fundtitle;
    }

    public void setFundtitle(String fundtitle) {
        this.fundtitle = fundtitle;
    }

    public String getSrcfileno() {
        return srcfileno;
    }

    public void setSrcfileno(String srcfileno) {
        this.srcfileno = srcfileno;
    }

    @Override
    public String toString() {
        return "CopyInfo{" +
                "packnumcek='" + packnumcek + '\'' +
                ", fundid='" + fundid + '\'' +
                ", pageno='" + pageno + '\'' +
                ", descsheetname='" + descsheetname + '\'' +
                ", copycount='" + copycount + '\'' +
                ", fundtitle='" + fundtitle + '\'' +
                ", srcfileno='" + srcfileno + '\'' +
                ", baoyou='" + baoyou + '\'' +
                ", quanti='" + quanti + '\'' +
                ", beginrow='" + beginrow + '\'' +
                ", endrow='" + endrow + '\'' +
                '}';
    }

    public String getBaoyou() {
        return baoyou;
    }

    public void setBaoyou(String baoyou) {
        this.baoyou = baoyou;
    }

    public String getQuanti() {
        return quanti;
    }

    public void setQuanti(String quanti) {
        this.quanti = quanti;
    }

    public boolean isCopy() {
        return this.pageno.equals("COPY");
    }

    public boolean isByCopy() {
        return this.baoyou.equals("COPY");
    }

    public boolean isQtCopy() {
        return this.quanti.equals("COPY");
    }

    public boolean isGuonei() {
        return this.ruliobject.contains("国内");
    }

    public int getBeginrow() {
        return beginrow;
    }

    public void setBeginrow(int beginrow) {
        this.beginrow = beginrow;
    }

    public int getEndrow() {
        return endrow;
    }

    public void setEndrow(int endrow) {
        this.endrow = endrow;
    }

    public int getRealcopycount() {
        return realcopycount;
    }

    public void setRealcopycount(int realcopycount) {
        this.realcopycount = realcopycount;
    }
}
