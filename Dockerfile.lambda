FROM public.ecr.aws/lambda/python:latest

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt --target "${LAMBDA_TASK_ROOT}"

COPY ./src/ ${LAMBDA_TASK_ROOT}

CMD [ "legoParser.handler" ]
