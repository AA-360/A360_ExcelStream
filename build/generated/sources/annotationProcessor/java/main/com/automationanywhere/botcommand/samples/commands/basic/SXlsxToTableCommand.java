package com.automationanywhere.botcommand.samples.commands.basic;

import com.automationanywhere.bot.service.GlobalSessionContext;
import com.automationanywhere.botcommand.BotCommand;
import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.commandsdk.i18n.Messages;
import com.automationanywhere.commandsdk.i18n.MessagesFactory;
import java.lang.Boolean;
import java.lang.ClassCastException;
import java.lang.Deprecated;
import java.lang.Double;
import java.lang.NullPointerException;
import java.lang.Object;
import java.lang.String;
import java.lang.Throwable;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;
import java.util.Optional;
import java.util.stream.Collectors;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

public final class SXlsxToTableCommand implements BotCommand {
  private static final Logger logger = LogManager.getLogger(SXlsxToTableCommand.class);

  private static final Messages MESSAGES_GENERIC = MessagesFactory.getMessages("com.automationanywhere.commandsdk.generic.messages");

  @Deprecated
  public Optional<Value> execute(Map<String, Value> parameters, Map<String, Object> sessionMap) {
    return execute(null, parameters, sessionMap);
  }

  public Optional<Value> execute(GlobalSessionContext globalSessionContext,
      Map<String, Value> parameters, Map<String, Object> sessionMap) {
    logger.traceEntry(() -> parameters != null ? parameters.entrySet().stream().filter(en -> !Arrays.asList( new String[] {}).contains(en.getKey()) && en.getValue() != null).collect(Collectors.toMap(Map.Entry::getKey, Map.Entry::getValue)).toString() : null, ()-> sessionMap != null ?sessionMap.toString() : null);
    SXlsxToTable command = new SXlsxToTable();
    HashMap<String, Object> convertedParameters = new HashMap<String, Object>();
    if(parameters.containsKey("file") && parameters.get("file") != null && parameters.get("file").get() != null) {
      convertedParameters.put("file", parameters.get("file").get());
      if(convertedParameters.get("file") !=null && !(convertedParameters.get("file") instanceof String)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","file", "String", parameters.get("file").get().getClass().getSimpleName()));
      }
    }
    if(convertedParameters.get("file") == null) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","file"));
    }
    if(convertedParameters.containsKey("file")) {
      String filePath= ((String)convertedParameters.get("file"));
      int lastIndxDot = filePath.lastIndexOf(".");
      if (lastIndxDot == -1 || lastIndxDot >= filePath.length()) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.FileExtension","file","xlsx"));
      }
      String fileExtension = filePath.substring(lastIndxDot + 1);
      if(!Arrays.stream("xlsx".split(",")).anyMatch(fileExtension::equalsIgnoreCase))  {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.FileExtension","file","xlsx"));
      }

    }
    if(parameters.containsKey("getSheetBy") && parameters.get("getSheetBy") != null && parameters.get("getSheetBy").get() != null) {
      convertedParameters.put("getSheetBy", parameters.get("getSheetBy").get());
      if(convertedParameters.get("getSheetBy") !=null && !(convertedParameters.get("getSheetBy") instanceof String)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","getSheetBy", "String", parameters.get("getSheetBy").get().getClass().getSimpleName()));
      }
    }
    if(convertedParameters.get("getSheetBy") == null) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","getSheetBy"));
    }
    if(convertedParameters.get("getSheetBy") != null) {
      switch((String)convertedParameters.get("getSheetBy")) {
        case "name" : {
          if(parameters.containsKey("sheetName") && parameters.get("sheetName") != null && parameters.get("sheetName").get() != null) {
            convertedParameters.put("sheetName", parameters.get("sheetName").get());
            if(convertedParameters.get("sheetName") !=null && !(convertedParameters.get("sheetName") instanceof String)) {
              throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","sheetName", "String", parameters.get("sheetName").get().getClass().getSimpleName()));
            }
          }
          if(convertedParameters.get("sheetName") == null) {
            throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","sheetName"));
          }


        } break;
        case "index" : {
          if(parameters.containsKey("sheetIndex") && parameters.get("sheetIndex") != null && parameters.get("sheetIndex").get() != null) {
            convertedParameters.put("sheetIndex", parameters.get("sheetIndex").get());
            if(convertedParameters.get("sheetIndex") !=null && !(convertedParameters.get("sheetIndex") instanceof Double)) {
              throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","sheetIndex", "Double", parameters.get("sheetIndex").get().getClass().getSimpleName()));
            }
          }
          if(convertedParameters.get("sheetIndex") == null) {
            throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","sheetIndex"));
          }
          if(convertedParameters.containsKey("sheetIndex")) {
            try {
              if(convertedParameters.get("sheetIndex") != null && !((double)convertedParameters.get("sheetIndex") >= 0)) {
                throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.GreaterThanEqualTo","sheetIndex", "0"));
              }
            }
            catch(ClassCastException e) {
              throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","sheetIndex", "Number", convertedParameters.get("sheetIndex").getClass().getSimpleName()));
            }
            catch(NullPointerException e) {
              throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","sheetIndex"));
            }

          }

        } break;
        default : throw new BotCommandException(MESSAGES_GENERIC.getString("generic.InvalidOption","getSheetBy"));
      }
    }

    if(parameters.containsKey("Columns") && parameters.get("Columns") != null && parameters.get("Columns").get() != null) {
      convertedParameters.put("Columns", parameters.get("Columns").get());
      if(convertedParameters.get("Columns") !=null && !(convertedParameters.get("Columns") instanceof String)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","Columns", "String", parameters.get("Columns").get().getClass().getSimpleName()));
      }
    }
    if(convertedParameters.get("Columns") == null) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","Columns"));
    }

    if(parameters.containsKey("hasHeaders") && parameters.get("hasHeaders") != null && parameters.get("hasHeaders").get() != null) {
      convertedParameters.put("hasHeaders", parameters.get("hasHeaders").get());
      if(convertedParameters.get("hasHeaders") !=null && !(convertedParameters.get("hasHeaders") instanceof Boolean)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","hasHeaders", "Boolean", parameters.get("hasHeaders").get().getClass().getSimpleName()));
      }
    }
    if(convertedParameters.get("hasHeaders") == null) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","hasHeaders"));
    }

    if(parameters.containsKey("RowStartCheck") && parameters.get("RowStartCheck") != null && parameters.get("RowStartCheck").get() != null) {
      convertedParameters.put("RowStartCheck", parameters.get("RowStartCheck").get());
      if(convertedParameters.get("RowStartCheck") !=null && !(convertedParameters.get("RowStartCheck") instanceof Boolean)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","RowStartCheck", "Boolean", parameters.get("RowStartCheck").get().getClass().getSimpleName()));
      }
    }
    if(convertedParameters.get("RowStartCheck") == null) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","RowStartCheck"));
    }
    if(convertedParameters.get("RowStartCheck") != null && (Boolean)convertedParameters.get("RowStartCheck")) {
      if(parameters.containsKey("RowStart") && parameters.get("RowStart") != null && parameters.get("RowStart").get() != null) {
        convertedParameters.put("RowStart", parameters.get("RowStart").get());
        if(convertedParameters.get("RowStart") !=null && !(convertedParameters.get("RowStart") instanceof Double)) {
          throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","RowStart", "Double", parameters.get("RowStart").get().getClass().getSimpleName()));
        }
      }
      if(convertedParameters.get("RowStart") == null) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","RowStart"));
      }
      if(convertedParameters.containsKey("RowStart")) {
        try {
          if(convertedParameters.get("RowStart") != null && !((double)convertedParameters.get("RowStart") >= 1)) {
            throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.GreaterThanEqualTo","RowStart", "1"));
          }
        }
        catch(ClassCastException e) {
          throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","RowStart", "Number", convertedParameters.get("RowStart").getClass().getSimpleName()));
        }
        catch(NullPointerException e) {
          throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","RowStart"));
        }

      }
    }

    if(parameters.containsKey("Decimal") && parameters.get("Decimal") != null && parameters.get("Decimal").get() != null) {
      convertedParameters.put("Decimal", parameters.get("Decimal").get());
      if(convertedParameters.get("Decimal") !=null && !(convertedParameters.get("Decimal") instanceof Boolean)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","Decimal", "Boolean", parameters.get("Decimal").get().getClass().getSimpleName()));
      }
    }
    if(convertedParameters.get("Decimal") == null) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","Decimal"));
    }
    if(convertedParameters.get("Decimal") != null && (Boolean)convertedParameters.get("Decimal")) {
      if(parameters.containsKey("DecimalPoints") && parameters.get("DecimalPoints") != null && parameters.get("DecimalPoints").get() != null) {
        convertedParameters.put("DecimalPoints", parameters.get("DecimalPoints").get());
        if(convertedParameters.get("DecimalPoints") !=null && !(convertedParameters.get("DecimalPoints") instanceof Double)) {
          throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","DecimalPoints", "Double", parameters.get("DecimalPoints").get().getClass().getSimpleName()));
        }
      }
      if(convertedParameters.get("DecimalPoints") == null) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","DecimalPoints"));
      }
      if(convertedParameters.containsKey("DecimalPoints")) {
        try {
          if(convertedParameters.get("DecimalPoints") != null && !((double)convertedParameters.get("DecimalPoints") >= 0)) {
            throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.GreaterThanEqualTo","DecimalPoints", "0"));
          }
        }
        catch(ClassCastException e) {
          throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","DecimalPoints", "Number", convertedParameters.get("DecimalPoints").getClass().getSimpleName()));
        }
        catch(NullPointerException e) {
          throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","DecimalPoints"));
        }

      }
    }

    try {
      Optional<Value> result =  Optional.ofNullable(command.action((String)convertedParameters.get("file"),(String)convertedParameters.get("getSheetBy"),(String)convertedParameters.get("sheetName"),(Double)convertedParameters.get("sheetIndex"),(String)convertedParameters.get("Columns"),(Boolean)convertedParameters.get("hasHeaders"),(Boolean)convertedParameters.get("RowStartCheck"),(Double)convertedParameters.get("RowStart"),(Boolean)convertedParameters.get("Decimal"),(Double)convertedParameters.get("DecimalPoints")));
      return logger.traceExit(result);
    }
    catch (ClassCastException e) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.IllegalParameters","action"));
    }
    catch (BotCommandException e) {
      logger.fatal(e.getMessage(),e);
      throw e;
    }
    catch (Throwable e) {
      logger.fatal(e.getMessage(),e);
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.NotBotCommandException",e.getMessage()),e);
    }
  }
}
