package com.beingidly.litexl;

import com.beingidly.litexl.crypto.SheetHasher;
import com.beingidly.litexl.crypto.SheetHasher.SheetProtectionInfo;
import com.beingidly.litexl.crypto.WriteProtection;
import org.jspecify.annotations.Nullable;

import java.util.Arrays;

/**
 * Manages write protection (file sharing) settings for a workbook.
 *
 * <p>This class handles the {@code fileSharing} element in workbook.xml,
 * which recommends opening the file as read-only and optionally requires
 * a password to modify the file.</p>
 *
 * <p>Passwords are accepted as char[] and immediately hashed for security -
 * the password itself is never stored.</p>
 *
 * <p>This class is <b>not thread-safe</b>.
 * External synchronization is required for concurrent access.</p>
 */
public final class WriteProtectionManager {

    private static final SheetHasher HASHER = new SheetHasher();

    private @Nullable WriteProtection protection;
    private @Nullable SheetProtectionInfo passwordInfo;

    WriteProtectionManager() {
    }

    /**
     * Sets write protection with a password.
     *
     * <p>The password is immediately hashed and the char array is cleared.
     * The password itself is never stored.</p>
     *
     * @param password the password (will be cleared after hashing), or null for no password
     * @param userName the name of the user setting the protection
     */
    public void protect(char @Nullable [] password, String userName) {
        this.protection = new WriteProtection(true, userName);
        if (password != null && password.length > 0) {
            try {
                this.passwordInfo = HASHER.hash(new String(password));
            } finally {
                Arrays.fill(password, '\0');
            }
        } else {
            this.passwordInfo = null;
        }
    }

    /**
     * Sets write protection without a password (read-only recommended).
     *
     * @param userName the name of the user setting the protection
     */
    public void protect(String userName) {
        this.protection = new WriteProtection(true, userName);
        this.passwordInfo = null;
    }

    /**
     * Removes write protection.
     */
    public void remove() {
        this.protection = null;
        this.passwordInfo = null;
    }

    /**
     * Returns the write protection settings, or null if not set.
     */
    public @Nullable WriteProtection protection() {
        return protection;
    }

    /**
     * Returns true if write protection is set.
     */
    public boolean isProtected() {
        return protection != null;
    }

    /**
     * Returns the password info for internal use during save.
     * Returns null if no password was set.
     */
    @Nullable SheetProtectionInfo passwordInfo() {
        return passwordInfo;
    }

    /**
     * Sets the password info directly (for internal use when reading).
     */
    void setPasswordInfo(@Nullable SheetProtectionInfo info) {
        this.passwordInfo = info;
    }

    /**
     * Sets the protection directly (for internal use when reading).
     */
    void setProtection(@Nullable WriteProtection protection) {
        this.protection = protection;
    }
}
