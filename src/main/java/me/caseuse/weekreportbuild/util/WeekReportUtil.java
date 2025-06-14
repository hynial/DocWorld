package me.caseuse.weekreportbuild.util;

import me.caseuse.weekreportbuild.entiry.CoordinateEntity;

import java.io.File;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class WeekReportUtil {
    public static Map<String, String> getTemplateValueMap(List<String> vars, List<String> contentLines) {
        Map<String, String> varValueMap = new HashMap<>();
        boolean keyFlag = false;
        String k = "";
        String v = "";
        for (String cLine : contentLines) {
//            String line = cLine.trim();
            String line = cLine;
            if (vars.contains(line)) { // template variable
                if (!keyFlag) {
                    keyFlag = true;
                } else {
                    varValueMap.put(k, v.endsWith("\n") ? v.substring(0, v.length() - 1) : v);
                    v = "";
                }
                k = line;
            } else {
                v += line + "\n";
            }
        }

        if (!varValueMap.containsKey(k)) {
            varValueMap.put(k, v.endsWith("\n") ? v.substring(0, v.length() - 1) : v);
        }

        return varValueMap;
    }

    public static List<String> readAllVariables(String varPath) throws IOException {
        File variablesFile = new File(varPath);
        List<String> varLines = Files.readAllLines(variablesFile.toPath(), StandardCharsets.UTF_8);
        List<String> vars = new ArrayList<>();
        for (String var : varLines) {
            if (var == null) continue;;
            if (var.trim().isEmpty()) continue;
            vars.add(var.trim());
        }

        return vars;
    }

    public static Map<String, List<CoordinateEntity>> loadAllVarCoordinate(String varPath) throws IOException {
        File variablesFile = new File(varPath);
        List<String> varLines = Files.readAllLines(variablesFile.toPath(), StandardCharsets.UTF_8);
        Map<String, List<CoordinateEntity>> map = new HashMap<>();
        for (String var : varLines) {
            if (var == null) continue;;
            if (var.trim().isEmpty()) continue;

            String[] splits = var.trim().split("\\|");
            if (splits.length != 2) {
                throw new RuntimeException("VarCoordinateLengthError:" + splits.length + ", For:" + var);
            }

            String coordinates = splits[1];
            String[] coors = coordinates.trim().split("!");
            List<CoordinateEntity> coordinateEntities = new ArrayList<>();
            for (String coordinate : coors) {
                String[] cs = coordinate.trim().split(",");
                if (cs.length != 2) {
                    throw new RuntimeException("VarCoordinateLengthError:" + splits.length + ", For:" + var + ",Coor:" + coordinate);
                }

                coordinateEntities.add(new CoordinateEntity(Integer.parseInt(cs[0].trim()), Integer.parseInt(cs[1].trim())));
            }

            map.put(splits[0].trim(), coordinateEntities);
        }

        return map;
    }
}
