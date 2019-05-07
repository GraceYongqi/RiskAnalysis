#!/bin/bash
i=1
for file in $(ls *.docx); do
#  echo ${file}
  mv ${file} ${i}.docx
  ((i=$i+1))
done
