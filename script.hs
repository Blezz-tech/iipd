main :: IO ()
main = putStrLn $ concatMap task [34..59]
  where
    task :: Int -> String
    task x = "## Задача " ++ show x ++ "\n\n\
      \### Условие задачи\n\n\n\n\
      \### Решение задачи\n\n\n\n\
      \### Ответ\n\n\n\n\
      \### Проверка\n\n\n\n"