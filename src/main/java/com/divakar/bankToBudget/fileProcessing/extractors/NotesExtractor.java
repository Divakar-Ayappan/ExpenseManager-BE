package com.divakar.bankToBudget.fileProcessing.extractors;

import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class NotesExtractor {

    public static List<String> getNotesFromTransactionsRorTMB(final List<String> transactions) {

        Pattern pattern = Pattern.compile("^UPI/[^/]+/[^/]+/[^/]+/([^/]+)");

        List<String> notes = new ArrayList<>();
        for (String txn : transactions) {
            Matcher matcher = pattern.matcher(txn);
            if (matcher.find()) {
                String note = matcher.group(1).trim();
                notes.add(note);
            } else {
                notes.add("Unknown");
            }
        }
        return notes;
    }

}