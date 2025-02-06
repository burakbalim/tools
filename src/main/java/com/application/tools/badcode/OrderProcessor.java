package com.application.tools.badcode;


import java.util.List;

public class OrderProcessor {

    public void process(List<String> ord) {
        double total = 0;
        for (String o : ord) {
            String[] parts = o.split(",");
            double price = Double.parseDouble(parts[1]);
            int qty = Integer.parseInt(parts[2]);
            total += price * qty;

            // Eğer sipariş tutarı 1000 TL üzerindeyse indirim uygula
            if (total > 1000) {
                total -= total * 0.1; // Magic number
            }
        }
        System.out.println("Total Amount: " + total);
    }
}
