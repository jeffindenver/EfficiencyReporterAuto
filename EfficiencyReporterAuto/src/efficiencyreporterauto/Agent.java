package efficiencyreporterauto;

import java.time.Duration;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.time.temporal.ChronoUnit;
import java.util.StringJoiner;

/**
 *
 * @author JShepherd
 */
public class Agent {

    public static Duration getTimeAsDuration(LocalTime time) {
        Duration seconds = Duration.of(time.toSecondOfDay(), ChronoUnit.SECONDS);
        return seconds;
    }

    private final String lastName;
    private final DateTimeFormatter timeFormatter;

    private Duration loginTime;
    private Duration workingTime;
    private Duration talkTime;
    private Duration acwTime;

    public Agent(String fullName) {
        timeFormatter = DateTimeFormatter.ofPattern("HH:mm:ss");
        this.lastName = fullName;

        loginTime = Duration.ZERO;
        workingTime = Duration.ZERO;
        talkTime = Duration.ZERO;
        acwTime = Duration.ZERO;
    }

    public String getLastName() {
        return lastName;
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
    public String toString() {
        StringJoiner joiner = new StringJoiner(",");

        double dLoginTime = loginTime.toMillis() / 1000;
        double dWorkingTime = workingTime.toMillis() / 1000;
        double dTalkTime = talkTime.toMillis() / 1000;
        double dAcwTime = acwTime.toMillis() / 1000;

        joiner.add(this.lastName);
        joiner.add(String.valueOf(dLoginTime / 60));
        joiner.add(String.valueOf(dWorkingTime / 60));
        joiner.add(String.valueOf(dTalkTime / 60));
        joiner.add(String.valueOf(dAcwTime / 60));

        return joiner.toString();
    }

}
