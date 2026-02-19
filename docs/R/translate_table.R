translate_table <- function(df, equiv_tbl){
  
  variables <- unique(equiv_tbl$VARIABLE)
  col_names <- colnames(df)
  for (variable in variables) {
    if (!(variable %in% col_names)) next
    
    # debugging
    print(variable)
    
    tmp <- equiv_tbl %>%
      filter(VARIABLE == variable) %>%
      select(CODE, EQUIV)
    
    df <- df %>%
      mutate(!!paste0(variable, "_int") := !!sym(variable)) %>%
      rename(temp = variable) %>%
      left_join(tmp, by = c("temp" = "CODE")) %>%
      mutate(temp = ifelse(is.na(EQUIV), temp, EQUIV)) %>%
      select(-EQUIV) %>%
      rename(!!variable := temp) %>%
      relocate(!!paste0(variable, "_int"), .before = !!sym(variable))
  }
  return(df)
}
