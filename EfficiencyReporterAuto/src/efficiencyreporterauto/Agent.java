package efficiencyreporterauto;

import java.time.Duration;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.time.temporal.ChronoUnit;
import java.util.Comparator;
import java.util.StringJoiner;

/**
 *
 * @author JShepherd
 */
public class Agent implements Comparable<Agent> {
    public final static Comparator<Agent> lnameComparator = new Comparator<Agent>() {
        @Override
        public int compare(Agent t, Agent t1) {
            return t.getLname().compareTo(t1.getLname());
        }
        
        @Override
        public Comparator<Agent> reversed() {
            return Comparator.super.reversed(); 
        }
    };

    public static Duration getTimeAsDuration(LocalTime time) {
        Duration seconds = Duration.of(time.toSecondOfDay(), ChronoUnit.SECONDS);
        return seconds;
    }

    private String fname;
    private String lname;
    private final String userID;
    private final DateTimeFormatter timeFormatter;

    private Duration loginTime;
    private Duration workingTime;
    private Duration talkTime;
    private Duration acwTime;

    public Agent(String userID) {
        timeFormatter = DateTimeFormatter.ofPattern("HH:mm:ss");
        this.userID = userID;
        this.fname = "";
        this.lname = "";
        loginTime = Duration.ZERO;
        workingTime = Duration.ZERO;
        talkTime = Duration.ZERO;
        acwTime = Duration.ZERO;
    }

    public String getUserID() {
        return userID;
    }

    public String getFullname() {
        return fname + " " + lname;
    }
   
    public String getFname() {
        return fname;
    }

    void setFname(String fname) {
        this.fname = fname;
    }

    public String getLname() {
        return lname;
    }

    void setLname(String lname) {
        this.lname = lname;
    }

    public DateTimeFormatter getTimeFormatter() {
        return timeFormatter;
    }
    
    public Duration getLoginTime() {
        return loginTime;
    }

    public void addLoginTime(String someTime) {
        LocalTime additionalTime = parseTime(someTime);
        this.loginTime = loginTime.plus(Agent.getTimeAsDuration(additionalTime));
    }

    public Duration getWorkingTime() {
        return workingTime;
    }

    public void addWorkingTime(String someTime) {
        LocalTime additionalTime = parseTime(someTime);
        this.workingTime = workingTime.plus(Agent.getTimeAsDuration(additionalTime));
    }

    public Duration getTalkTime() {
        return talkTime;
    }

    public void addTalkTime(String someTime) {
        LocalTime additionalTime = parseTime(someTime);
        this.talkTime = talkTime.plus(Agent.getTimeAsDuration(additionalTime));
    }

    public Duration getAcwTime() {
        return acwTime;
    }

    public void addAcwTime(String someTime) {
        LocalTime additionalTime = parseTime(someTime);
        this.acwTime = acwTime.plus(Agent.getTimeAsDuration(additionalTime));
    }

    private LocalTime parseTime(String someTime) {
        LocalTime additionalTime = LocalTime.MIN;
        try {
            additionalTime = LocalTime.parse(someTime, timeFormatter);
        } catch (DateTimeParseException e) {
            System.out.println(e.getMessage());
            return additionalTime;
        }
        return additionalTime;
    }
    
    @Override
    public int compareTo(Agent t) {
        return this.getLname().compareToIgnoreCase(t.getLname());
    }
    
    @Override
    public String toString() {
        StringJoiner joiner = new StringJoiner(",");

        double dLoginTime = getLoginTime().toMillis() / 1000;
        double dWorkingTime = getWorkingTime().toMillis() / 1000;
        double dTalkTime = getTalkTime().toMillis() / 1000;
        double dAcwTime = getAcwTime().toMillis() / 1000;
        String fullname = getFname() + " " + getLname();

        joiner.add(fullname);
        joiner.add(String.valueOf(dLoginTime / 60));
        joiner.add(String.valueOf(dWorkingTime / 60));
        joiner.add(String.valueOf(dTalkTime / 60));
        joiner.add(String.valueOf(dAcwTime / 60));

        return joiner.toString();
    }

}
