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
    final static Comparator<Agent> lnameComparator = new Comparator<Agent>() {
        @Override
        public int compare(Agent t, Agent t1) {
            return t.getLname().compareTo(t1.getLname());
        }
        
        @Override
        public Comparator<Agent> reversed() {
            return Comparator.super.reversed(); 
        }
    };

    private static Duration getTimeAsDuration(LocalTime time) {
       return Duration.of(time.toSecondOfDay(), ChronoUnit.SECONDS);
    }

    private String fname;
    private String lname;
    private final String userID;
    private final DateTimeFormatter timeFormatter;

    private Duration loginTime;
    private Duration workingTime;
    private Duration talkTime;
    private Duration acwTime;
    private Duration breaksAndOtherTime;

    public Agent(String userID) {
        timeFormatter = DateTimeFormatter.ofPattern("HH:mm:ss");
        this.userID = userID;
        this.fname = "";
        this.lname = "";
        loginTime = Duration.ZERO;
        workingTime = Duration.ZERO;
        talkTime = Duration.ZERO;
        acwTime = Duration.ZERO;
        breaksAndOtherTime = Duration.ZERO;
    }

    String getUserID() {
        return userID;
    }

    String getFullname() {
        return fname + " " + lname;
    }
   
    String getFname() {
        return fname;
    }

    void setFname(String fname) {
        this.fname = fname;
    }

    private String getLname() {
        return lname;
    }

    void setLname(String lname) {
        this.lname = lname;
    }

    Duration getLoginTime() {
        return loginTime;
    }

    void addLoginTime(String someTime) {
        LocalTime additionalTime = parseTime(someTime);
        this.loginTime = loginTime.plus(Agent.getTimeAsDuration(additionalTime));
    }

    Duration getWorkingTime() {
        return workingTime;
    }

    void addWorkingTime(String someTime) {
        LocalTime additionalTime = parseTime(someTime);
        this.workingTime = workingTime.plus(Agent.getTimeAsDuration(additionalTime));
    }

    Duration getTalkTime() {
        return talkTime;
    }

    void addTalkTime(String someTime) {
        LocalTime additionalTime = parseTime(someTime);
        this.talkTime = talkTime.plus(Agent.getTimeAsDuration(additionalTime));
    }

    Duration getAcwTime() {
        return acwTime;
    }

    void addAcwTime(String someTime) {
        LocalTime additionalTime = parseTime(someTime);
        this.acwTime = acwTime.plus(Agent.getTimeAsDuration(additionalTime));
    }
    
    void addBreaksAndOtherTime(String someTime) {
        LocalTime additionalTime = parseTime(someTime);
        this.breaksAndOtherTime = breaksAndOtherTime.plus(Agent.getTimeAsDuration(additionalTime));
    }

    Duration getBreaksAndOtherTime() {
        return breaksAndOtherTime;
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

        double dLoginTime = getLoginTime().toMillis() / 1000.0;
        double dWorkingTime = getWorkingTime().toMillis() / 1000.0;
        double dTalkTime = getTalkTime().toMillis() / 1000.0;
        double dAcwTime = getAcwTime().toMillis() / 1000.0;
        double dBreaksAndOtherTime = getBreaksAndOtherTime().toMillis() / 1000.0;
        String fullname = getFname() + " " + getLname();

        joiner.add(fullname);
        joiner.add(String.valueOf(dLoginTime / 60));
        joiner.add(String.valueOf(dWorkingTime / 60));
        joiner.add(String.valueOf(dTalkTime / 60));
        joiner.add(String.valueOf(dAcwTime / 60));
        joiner.add(String.valueOf(dBreaksAndOtherTime / 60));

        return joiner.toString();
    }

}
