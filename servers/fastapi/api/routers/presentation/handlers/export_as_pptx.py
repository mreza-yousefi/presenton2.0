import os
import uuid
from api.models import LogMetadata
from api.routers.presentation.mixins.fetch_presentation_assets import (
    FetchPresentationAssetsMixin,
)
from api.routers.presentation.models import (
    ExportAsRequest,
    PresentationAndPath,
)
from api.services.logging import LoggingService
from api.services.instances import TEMP_FILE_SERVICE
from api.sql_models import PresentationSqlModel
from api.utils.utils import get_presentation_dir, sanitize_filename
from ppt_generator.pptx_presentation_creator import PptxPresentationCreator
from api.services.database import get_sql_session


class ExportAsPptxHandler(FetchPresentationAssetsMixin):

    def __init__(self, data: ExportAsRequest):
        self.data = data
        self.template_file_path: Optional[str] = None

        self.session = str(uuid.uuid4())
        self.temp_dir = TEMP_FILE_SERVICE.create_temp_dir(self.session)

        self.presentation_dir = get_presentation_dir(self.data.presentation_id)

    def __del__(self):
        TEMP_FILE_SERVICE.cleanup_temp_dir(self.temp_dir)

    async def post(self, logging_service: LoggingService, log_metadata: LogMetadata):
        # Log data excluding file content for brevity
        loggable_data = self.data.model_dump(mode="json", exclude={'template_file'})
        if self.data.template_file:
            loggable_data['template_file_name'] = self.data.template_file.filename
        logging_service.logger.info(
            logging_service.message(loggable_data),
            extra=log_metadata.model_dump(),
        )

        template_path_for_creator = None
        if self.data.template_file:
            try:
                # Save uploaded template to a temporary file
                self.template_file_path = os.path.join(self.temp_dir, f"user_template_{uuid.uuid4()}.pptx")
                with open(self.template_file_path, "wb") as buffer:
                    buffer.write(await self.data.template_file.read())
                template_path_for_creator = self.template_file_path
                logging_service.logger.info(f"User template saved to: {template_path_for_creator}")
            except Exception as e:
                logging_service.logger.error(f"Error saving uploaded template: {e}", extra=log_metadata.model_dump())
                # Optionally, decide if to proceed without template or raise error
                # For now, proceeding without template if save fails

        await self.fetch_presentation_assets()

        with get_sql_session() as sql_session:
            presentation = sql_session.get(
                PresentationSqlModel, self.data.presentation_id
            )

        ppt_path = os.path.join(
            self.presentation_dir,
            sanitize_filename(f"{presentation.title}.pptx")
        )
        ppt_creator = PptxPresentationCreator(
            self.data.pptx_model,
            self.temp_dir,
            template_path=template_path_for_creator
        )
        ppt_creator.create_ppt()
        ppt_creator.save(ppt_path)

        response = PresentationAndPath(
            presentation_id=self.data.presentation_id, path=ppt_path
        )

        with get_sql_session() as sql_session:
            presentation = sql_session.get(
                PresentationSqlModel, self.data.presentation_id
            )
            presentation.file = ppt_path
            sql_session.commit()

        logging_service.logger.info(
            logging_service.message(response.model_dump(mode="json")),
            extra=log_metadata.model_dump(),
        )

        return response
