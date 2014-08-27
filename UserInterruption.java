
import java.io.Serializable;

/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

/**
 *
 * @author doyin
 */
public class UserInterruption extends Exception implements Serializable {
    public String getMessage(){
        return "UserInterruption";
    }
}
